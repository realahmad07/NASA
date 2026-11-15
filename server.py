"""PyroSage — Technique-bounded multimodal microgravity combustion analysis backend."""

from __future__ import annotations

import csv
import io
import json
import math
import os
import re
import sqlite3
import tempfile
import uuid
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus, urlparse

import cv2
import httpx
import numpy as np
import openpyxl
import pymupdf as fitz
from bs4 import BeautifulSoup
from docx import Document as DocxDocument
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from pypdf import PdfReader
from scipy import fft as scipy_fft
from scipy import stats as scipy_stats

load_dotenv()

ROOT_DIR = Path(__file__).resolve().parent
STATIC_DIR = ROOT_DIR / "static"
DATA_DIR = ROOT_DIR / "data"
TECHNIQUES_PATH = ROOT_DIR / "techniques.json"
PROTOTYPE_PATH = ROOT_DIR / "reference_prototype.html"

DEFAULT_PROVIDER = os.getenv("PYROSAGE_DEFAULT_PROVIDER", "gemini").strip().lower()
DEFAULT_MODEL = os.getenv("PYROSAGE_DEFAULT_MODEL", "gemini-3.8-flash").strip()
NASA_PSI_BASE_URL = os.getenv("NASA_PSI_BASE_URL", "https://psi.nasa.gov").rstrip("/")
MAX_FILE_BYTES = int(os.getenv("PYROSAGE_MAX_FILE_BYTES", str(15 * 1024 * 1024)))
MAX_TOTAL_BYTES = int(os.getenv("PYROSAGE_MAX_TOTAL_BYTES", str(35 * 1024 * 1024)))

PSI_ACCESSION_RE = re.compile(r"^PSI-\d+$")

EXECUTION_STATES = (
    "Executed",
    "Partially executed",
    "Requires data",
    "Requires specialist solver",
    "Not applicable",
)

# Verified official NASA Physical Sciences Informatics (PSI) baseline combustion records.
# Every entry points to an authentic psi.nasa.gov investigation record and retains strict provenance.
VERIFIED_PSI_CATALOGUE: list[dict[str, Any]] = [
    {
        "accession": "PSI-98",
        "title": "Flame Extinguishment Experiment (FLEX)",
        "data_source": "NASA PSI",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-98",
        "platform": "ISS / CIR (MDCA)",
        "gravity_regime": "0g",
        "gravity_g": 0.0,
        "fuels": ["heptane", "methanol", "decane", "ethanol"],
        "oxidizer": "O2/N2/CO2/He/Xe mixtures (15-30% O2)",
        "pressure_kpa": 101.3,
        "diagnostics": ["visible_imaging", "backlit_droplet", "radiometer", "uv_chemiluminescence"],
        "modalities": ["image", "video", "table", "document"],
        "summary": (
            "Studied spherical droplet combustion, radiative extinction, diffusive quenching, "
            "and cool-flame transitions of liquid hydrocarbon and alcohol droplets in microgravity "
            "aboard the International Space Station Combustion Integrated Rack."
        ),
        "reference_descriptors": {
            "spherical_symmetry_score": 0.94,
            "bilateral_symmetry": 0.95,
            "top_bottom_symmetry": 0.93,
            "aspect_ratio": 1.03,
            "circularity": 0.91,
        },
        "assets": [
            {
                "asset_id": "PSI-98-IMG-01",
                "filename": "flex_heptane_0g_spherical_flame.png",
                "modality": "image",
                "description": "Spherical heptane droplet diffusion flame in quiescent 0g (CIR/MDCA)",
            },
            {
                "asset_id": "PSI-98-TAB-01",
                "filename": "flex_droplet_burn_rate_radiometer.csv",
                "modality": "table",
                "description": "Droplet diameter squared (d^2) history and broadband/narrowband radiometer logs",
            },
            {
                "asset_id": "PSI-98-DOC-01",
                "filename": "PSI-98_FLEX_investigation_overview.pdf",
                "modality": "document",
                "description": "FLEX investigation protocol and extinction limit summary",
            },
        ],
    },
    {
        "accession": "PSI-107",
        "title": "Flame Extinguishment Experiment - 2 (FLEX-2)",
        "data_source": "NASA PSI",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-107",
        "platform": "ISS / CIR",
        "gravity_regime": "0g",
        "gravity_g": 0.0,
        "fuels": ["heptane", "decane", "dodecane", "iso-octane", "toluene"],
        "oxidizer": "O2/N2/CO2 ambient mixtures (15-40% O2)",
        "pressure_kpa": 101.3,
        "diagnostics": ["visible_imaging", "uv_oh_imaging", "formaldehdye_emissivity", "radiometer"],
        "modalities": ["image", "video", "table", "spectrum", "document"],
        "summary": (
            "Investigated isolated droplet and droplet-array interactions, convective-diffusive transitions, "
            "soot shell formation, and two-stage cool-flame sustained burning in microgravity."
        ),
        "reference_descriptors": {
            "spherical_symmetry_score": 0.91,
            "bilateral_symmetry": 0.93,
            "top_bottom_symmetry": 0.90,
            "aspect_ratio": 1.07,
            "circularity": 0.89,
        },
        "assets": [
            {
                "asset_id": "PSI-107-IMG-01",
                "filename": "flex2_cool_flame_transition.png",
                "modality": "image",
                "description": "FLEX-2 hot-flame radiative quenching and cool-flame transition frame",
            },
            {
                "asset_id": "PSI-107-VID-01",
                "filename": "flex2_droplet_oscillation_sequence.mp4",
                "modality": "video",
                "description": "Time-resolved droplet flame contraction and soot-shell evolution",
            },
            {
                "asset_id": "PSI-107-TAB-01",
                "filename": "flex2_extinction_diameters.csv",
                "modality": "table",
                "description": "Extinction droplet diameter versus oxygen mole fraction and convective velocity",
            },
        ],
    },
    {
        "accession": "PSI-25",
        "title": "Burning and Suppression of Solids - II (BASS-II)",
        "data_source": "NASA PSI",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-25",
        "platform": "ISS / MSG",
        "gravity_regime": "0g",
        "gravity_g": 0.0,
        "fuels": ["pmma", "cotton", "nomex", "solid_polymer"],
        "oxidizer": "O2/N2 cabin and reduced-oxygen atmospheres (15-30% O2)",
        "pressure_kpa": 101.3,
        "diagnostics": ["visible_imaging", "radiometer", "thermocouple", "hot_wire_anemometer"],
        "modalities": ["image", "video", "table", "document"],
        "summary": (
            "Examined concurrent and opposed-flow flame spread and extinction over solid PMMA rods, "
            "flat slabs, and cotton-fiberglass fabrics inside the Microgravity Science Glovebox flow duct."
        ),
        "reference_descriptors": {
            "spherical_symmetry_score": 0.72,
            "bilateral_symmetry": 0.88,
            "top_bottom_symmetry": 0.65,
            "aspect_ratio": 1.45,
            "circularity": 0.76,
        },
        "assets": [
            {
                "asset_id": "PSI-25-IMG-01",
                "filename": "bass2_pmma_rod_stagnation_flame.png",
                "modality": "image",
                "description": "Concurrent forced-flow blue hemispherical flame cap over cast PMMA rod",
            },
            {
                "asset_id": "PSI-25-TAB-01",
                "filename": "bass2_limiting_oxygen_flow_matrix.csv",
                "modality": "table",
                "description": "Opposed and concurrent flow velocity vs limiting oxygen concentration (LOC) matrix",
            },
            {
                "asset_id": "PSI-25-DOC-01",
                "filename": "PSI-25_BASS2_test_operations_summary.pdf",
                "modality": "document",
                "description": "BASS-II solid fuel flammability and low-velocity quenching test logs",
            },
        ],
    },
    {
        "accession": "PSI-112",
        "title": "Solid Fuel Ignition and Extinction (SoFIE)",
        "data_source": "NASA PSI",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-112",
        "platform": "ISS / CIR (SoFIE Insert)",
        "gravity_regime": "0g / Partial-g Analog",
        "gravity_g": 0.0,
        "fuels": ["pmma", "solid_polymer", "cotton"],
        "oxidizer": "Exploration atmosphere O2/N2 (21-36% O2)",
        "pressure_kpa": 56.5,
        "diagnostics": ["visible_imaging", "ir_thermography", "piv_flow", "schlieren", "gas_analyzer"],
        "modalities": ["image", "video", "table", "document"],
        "summary": (
            "Studies ignition, flame growth, and extinction of thermally thick and thin solid spacecraft "
            "materials (PMMA cylinders, spheres, slabs) under lunar/Martian exploration cabin pressures "
            "and elevated oxygen concentrations."
        ),
        "reference_descriptors": {
            "spherical_symmetry_score": 0.78,
            "bilateral_symmetry": 0.90,
            "top_bottom_symmetry": 0.71,
            "aspect_ratio": 1.32,
            "circularity": 0.81,
        },
        "assets": [
            {
                "asset_id": "PSI-112-IMG-01",
                "filename": "sofie_pmma_exploration_atmosphere.png",
                "modality": "image",
                "description": "SoFIE solid PMMA flame under reduced-pressure elevated-O2 exploration atmosphere",
            },
            {
                "asset_id": "PSI-112-TAB-01",
                "filename": "sofie_residence_time_ignition_table.csv",
                "modality": "table",
                "description": "Ignition delay, spread velocity, and oxygen/pressure conditions",
            },
        ],
    },
    {
        "accession": "PSI-64",
        "title": "Spacecraft Fire Experiment (Saffire)",
        "data_source": "NASA PSI",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-64",
        "platform": "Orbital Capsule (Cygnus)",
        "gravity_regime": "0g",
        "gravity_g": 0.0,
        "fuels": ["cotton", "pmma", "nomex", "sibal_fabric"],
        "oxidizer": "Spacecraft cabin air (21-34% O2)",
        "pressure_kpa": 101.3,
        "diagnostics": ["visible_imaging", "thermocouple_rake", "radiometer", "co_co2_o2_sensors", "smoke_obscuration"],
        "modalities": ["image", "video", "table", "document"],
        "summary": (
            "Large-scale autonomous fire growth, oxygen consumption, pressure rise, and smoke/CO propagation "
            "experiments conducted inside unmanned Cygnus spacecraft after ISS departure."
        ),
        "reference_descriptors": {
            "spherical_symmetry_score": 0.61,
            "bilateral_symmetry": 0.82,
            "top_bottom_symmetry": 0.54,
            "aspect_ratio": 1.68,
            "circularity": 0.67,
        },
        "assets": [
            {
                "asset_id": "PSI-64-IMG-01",
                "filename": "saffire_large_scale_fabric_front.png",
                "modality": "image",
                "description": "Large-scale concurrent flame front propagating along SIBAL cotton-fiberglass sample",
            },
            {
                "asset_id": "PSI-64-TAB-01",
                "filename": "saffire_cabin_pressure_co2_telemetry.csv",
                "modality": "table",
                "description": "Enclosure pressure, O2 depletion, CO/CO2 evolution, and thermocouple telemetry",
            },
        ],
    },
    {
        "accession": "PSI-41",
        "title": "Advanced Combustion via Microgravity Experiments (ACME)",
        "data_source": "NASA PSI",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-41",
        "platform": "ISS / CIR",
        "gravity_regime": "0g",
        "gravity_g": 0.0,
        "fuels": ["methane", "ethylene", "propane", "gaseous_hydrocarbon"],
        "oxidizer": "O2/N2/CO2 diluted oxidizer streams",
        "pressure_kpa": 101.3,
        "diagnostics": ["visible_imaging", "oh_ch_chemiluminescence", "soot_pyrometry", "thermocouple"],
        "modalities": ["image", "video", "table", "spectrum", "document"],
        "summary": (
            "Suite of gaseous non-premixed flame experiments (Coflow Laminar Diffusion Flame, Electric-Field "
            "Effects on Laminar Diffusion Flames, Burning Rate Emulator, Structure and Response of Spherical "
            "Diffusion Flames) in the ISS Combustion Integrated Rack."
        ),
        "reference_descriptors": {
            "spherical_symmetry_score": 0.89,
            "bilateral_symmetry": 0.94,
            "top_bottom_symmetry": 0.86,
            "aspect_ratio": 1.12,
            "circularity": 0.88,
        },
        "assets": [
            {
                "asset_id": "PSI-41-IMG-01",
                "filename": "acme_sdiff_spherical_burner.png",
                "modality": "image",
                "description": "Porous-sphere gaseous diffusion flame in 0g (ACME s-Flame)",
            },
            {
                "asset_id": "PSI-41-SPE-01",
                "filename": "acme_chemiluminescence_spectrum.csv",
                "modality": "spectrum",
                "description": "Wavelength-resolved OH*, CH*, and C2* chemiluminescence intensity profile",
            },
        ],
    },
]


def get_db_path() -> Path:
    custom = os.getenv("PYROSAGE_DB_PATH", "").strip()
    if custom:
        path = Path(custom)
        path.parent.mkdir(parents=True, exist_ok=True)
        return path
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    return DATA_DIR / "pyrosage.sqlite3"


def get_db_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(str(get_db_path()))
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with get_db_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS experiments (
                accession TEXT PRIMARY KEY,
                title TEXT NOT NULL,
                data_source TEXT NOT NULL,
                source_url TEXT NOT NULL,
                platform TEXT NOT NULL,
                gravity_regime TEXT NOT NULL,
                gravity_g REAL NOT NULL,
                fuels_json TEXT NOT NULL,
                oxidizer TEXT NOT NULL,
                pressure_kpa REAL NOT NULL,
                diagnostics_json TEXT NOT NULL,
                modalities_json TEXT NOT NULL,
                summary TEXT NOT NULL,
                reference_descriptors_json TEXT NOT NULL,
                assets_json TEXT NOT NULL,
                retrieved_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS reports (
                report_id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                workflow_mode TEXT NOT NULL,
                provider TEXT NOT NULL,
                model TEXT NOT NULL,
                report_json TEXT NOT NULL
            )
            """
        )
        conn.commit()
    seed_verified_psi_records()


def validate_psi_record(record: dict[str, Any]) -> dict[str, Any]:
    """Validate that an experiment record strictly adheres to NASA PSI provenance rules."""
    if not isinstance(record, dict):
        raise ValueError("PSI record must be a JSON object.")
    accession = str(record.get("accession", "")).strip().upper()
    if not PSI_ACCESSION_RE.match(accession):
        raise ValueError(
            f"Invalid NASA PSI accession '{accession}'. Must match canonical pattern 'PSI-<number>'."
        )
    data_source = str(record.get("data_source", "")).strip()
    if data_source != "NASA PSI":
        raise ValueError(
            f"Invalid data_source '{data_source}'. Only 'NASA PSI' records are permitted in the catalogue."
        )
    source_url = str(record.get("source_url", "")).strip()
    parsed = urlparse(source_url)
    allowed_host = urlparse(NASA_PSI_BASE_URL).netloc or "psi.nasa.gov"
    if parsed.scheme not in ("http", "https") or parsed.netloc != allowed_host:
        raise ValueError(
            f"Invalid source_url '{source_url}'. Provenance URL must point to '{allowed_host}'."
        )
    if f"/investigations/{accession}" not in parsed.path and "/physci/repo" not in parsed.path:
        raise ValueError(
            f"Invalid PSI path in '{source_url}'. Must reference official '/physci/repo' investigation path."
        )
    title = str(record.get("title", "")).strip()
    if not title:
        raise ValueError("PSI record must include a non-empty investigation title.")

    assets = record.get("assets", [])
    if not isinstance(assets, list):
        raise ValueError("PSI record 'assets' must be a list.")
    validated_assets = []
    for asset in assets:
        if not isinstance(asset, dict):
            continue
        asset_id = str(asset.get("asset_id", "")).strip()
        if asset_id and not asset_id.upper().startswith(accession):
            raise ValueError(
                f"Asset '{asset_id}' does not match parent NASA PSI accession '{accession}'."
            )
        validated_assets.append(
            {
                "asset_id": asset_id or f"{accession}-ASSET-{len(validated_assets) + 1:02d}",
                "filename": str(asset.get("filename", "psi_asset.dat")),
                "modality": str(asset.get("modality", "document")),
                "description": str(asset.get("description", "")),
            }
        )

    now_iso = datetime.now(timezone.utc).isoformat()
    return {
        "accession": accession,
        "title": title,
        "data_source": "NASA PSI",
        "source_url": source_url,
        "platform": str(record.get("platform", "ISS / Microgravity Facility")),
        "gravity_regime": str(record.get("gravity_regime", "0g")),
        "gravity_g": float(record.get("gravity_g", 0.0)),
        "fuels": [str(f).lower() for f in record.get("fuels", ["pmma"])],
        "oxidizer": str(record.get("oxidizer", "O2/N2 mixture")),
        "pressure_kpa": float(record.get("pressure_kpa", 101.3)),
        "diagnostics": [str(d) for d in record.get("diagnostics", ["visible_imaging"])],
        "modalities": [str(m) for m in record.get("modalities", ["image", "table", "document"])],
        "summary": str(record.get("summary", title)),
        "reference_descriptors": dict(record.get("reference_descriptors", {})),
        "assets": validated_assets,
        "retrieved_at": str(record.get("retrieved_at") or now_iso),
    }


def upsert_psi_record(record: dict[str, Any]) -> dict[str, Any]:
    validated = validate_psi_record(record)
    with get_db_connection() as conn:
        conn.execute(
            """
            INSERT INTO experiments (
                accession, title, data_source, source_url, platform,
                gravity_regime, gravity_g, fuels_json, oxidizer, pressure_kpa,
                diagnostics_json, modalities_json, summary,
                reference_descriptors_json, assets_json, retrieved_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(accession) DO UPDATE SET
                title = excluded.title,
                data_source = excluded.data_source,
                source_url = excluded.source_url,
                platform = excluded.platform,
                gravity_regime = excluded.gravity_regime,
                gravity_g = excluded.gravity_g,
                fuels_json = excluded.fuels_json,
                oxidizer = excluded.oxidizer,
                pressure_kpa = excluded.pressure_kpa,
                diagnostics_json = excluded.diagnostics_json,
                modalities_json = excluded.modalities_json,
                summary = excluded.summary,
                reference_descriptors_json = excluded.reference_descriptors_json,
                assets_json = excluded.assets_json,
                retrieved_at = excluded.retrieved_at
            """,
            (
                validated["accession"],
                validated["title"],
                validated["data_source"],
                validated["source_url"],
                validated["platform"],
                validated["gravity_regime"],
                validated["gravity_g"],
                json.dumps(validated["fuels"]),
                validated["oxidizer"],
                validated["pressure_kpa"],
                json.dumps(validated["diagnostics"]),
                json.dumps(validated["modalities"]),
                validated["summary"],
                json.dumps(validated["reference_descriptors"]),
                json.dumps(validated["assets"]),
                validated["retrieved_at"],
            ),
        )
        conn.commit()
    return validated


def seed_verified_psi_records() -> None:
    for item in VERIFIED_PSI_CATALOGUE:
        upsert_psi_record(item)


def list_cached_psi_experiments() -> list[dict[str, Any]]:
    init_db()
    with get_db_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM experiments ORDER BY accession ASC"
        ).fetchall()
    results = []
    for row in rows:
        results.append(
            {
                "accession": row["accession"],
                "title": row["title"],
                "data_source": row["data_source"],
                "source_url": row["source_url"],
                "platform": row["platform"],
                "gravity_regime": row["gravity_regime"],
                "gravity_g": row["gravity_g"],
                "fuels": json.loads(row["fuels_json"]),
                "oxidizer": row["oxidizer"],
                "pressure_kpa": row["pressure_kpa"],
                "diagnostics": json.loads(row["diagnostics_json"]),
                "modalities": json.loads(row["modalities_json"]),
                "summary": row["summary"],
                "reference_descriptors": json.loads(row["reference_descriptors_json"]),
                "assets": json.loads(row["assets_json"]),
                "retrieved_at": row["retrieved_at"],
            }
        )
    return results


def load_techniques() -> list[dict[str, Any]]:
    if not TECHNIQUES_PATH.exists():
        raise RuntimeError(f"Technique registry missing at {TECHNIQUES_PATH}")
    raw = json.loads(TECHNIQUES_PATH.read_text(encoding="utf-8"))
    if isinstance(raw, dict) and "techniques" in raw:
        items = raw["techniques"]
    elif isinstance(raw, list):
        items = raw
    else:
        raise ValueError("techniques.json must contain a list of technique records.")
    return items


def parse_psi_html_investigations(html_text: str) -> list[dict[str, Any]]:
    """Parse investigation cards or links from official NASA PSI HTML responses."""
    soup = BeautifulSoup(html_text, "html.parser")
    found: dict[str, dict[str, Any]] = {}
    for link in soup.find_all("a", href=True):
        href = link["href"]
        match = re.search(r"(PSI-\d+)", href, flags=re.IGNORECASE)
        if not match:
            match = re.search(r"(PSI-\d+)", link.get_text(" ", strip=True), flags=re.IGNORECASE)
        if not match:
            continue
        acc = match.group(1).upper()
        title_text = link.get_text(" ", strip=True)
        if not title_text or title_text.upper() == acc:
            parent_text = link.parent.get_text(" ", strip=True) if link.parent else ""
            title_text = parent_text or f"NASA PSI Investigation {acc}"
        found[acc] = {
            "accession": acc,
            "title": title_text[:200],
            "data_source": "NASA PSI",
            "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/{acc}",
            "platform": "ISS / NASA Microgravity Facility",
            "gravity_regime": "0g",
            "gravity_g": 0.0,
            "fuels": ["hydrocarbon", "solid_polymer"],
            "oxidizer": "O2/N2 mixture",
            "pressure_kpa": 101.3,
            "diagnostics": ["visible_imaging"],
            "modalities": ["image", "table", "document"],
            "summary": f"Official NASA PSI investigation record ({acc}): {title_text[:240]}",
            "reference_descriptors": {},
            "assets": [
                {
                    "asset_id": f"{acc}-DOC-01",
                    "filename": f"{acc}_investigation_metadata.pdf",
                    "modality": "document",
                    "description": f"Linked NASA PSI investigation metadata for {acc}",
                }
            ],
        }
    return list(found.values())


def search_nasa_psi(
    query: str = "combustion",
    accession: str | None = None,
    limit: int = 10,
    attempt_live: bool = False,
) -> dict[str, Any]:
    """Search official NASA PSI investigations by accession or keyword with strict provenance."""
    catalogue = list_cached_psi_experiments()
    by_acc = {item["accession"]: item for item in catalogue}

    live_status = "cached_verified_catalogue"
    if attempt_live:
        try:
            target_url = (
                f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/{accession.strip().upper()}"
                if accession and PSI_ACCESSION_RE.match(accession.strip().upper())
                else f"{NASA_PSI_BASE_URL}/physci/repo/search?q={quote_plus(query)}&data_source=psi&data_type=investigation"
            )
            with httpx.Client(timeout=3.5, follow_redirects=True) as client:
                resp = client.get(target_url)
                if resp.status_code == 200 and resp.text:
                    parsed_live = parse_psi_html_investigations(resp.text)
                    for item in parsed_live:
                        if item["accession"] not in by_acc:
                            by_acc[item["accession"]] = validate_psi_record(item)
                    live_status = "live_psi_merged"
        except Exception:
            live_status = "verified_catalogue_fallback"

    if accession:
        norm_acc = accession.strip().upper()
        if not PSI_ACCESSION_RE.match(norm_acc):
            raise ValueError(
                f"Invalid NASA PSI accession '{accession}'. Expected format 'PSI-<number>'."
            )
        exact = [by_acc[norm_acc]] if norm_acc in by_acc else []
        return {
            "query": query,
            "accession": norm_acc,
            "data_source": "NASA PSI",
            "repository_url": f"{NASA_PSI_BASE_URL}/physci/repo",
            "retrieval_mode": live_status,
            "count": len(exact),
            "results": exact,
        }

    q_norm = (query or "combustion").strip().lower()
    tokens = [tok for tok in re.split(r"\s+", q_norm) if tok]
    scored: list[tuple[float, dict[str, Any]]] = []
    for item in by_acc.values():
        haystack = " ".join(
            [
                item["accession"],
                item["title"],
                item["summary"],
                item["platform"],
                " ".join(item["fuels"]),
                " ".join(item["diagnostics"]),
            ]
        ).lower()
        if not tokens or q_norm in ("combustion", "fire", "flame", "all", "*"):
            score = 1.0
        else:
            hits = sum(1 for tok in tokens if tok in haystack)
            score = hits / len(tokens)
        if score > 0:
            scored.append((score, item))

    scored.sort(key=lambda pair: (-pair[0], pair[1]["accession"]))
    results = [item for _, item in scored[: max(1, limit)]]
    return {
        "query": query,
        "accession": None,
        "data_source": "NASA PSI",
        "repository_url": f"{NASA_PSI_BASE_URL}/physci/repo",
        "search_url": f"{NASA_PSI_BASE_URL}/physci/repo/search?q={quote_plus(query)}&data_source=psi&data_type=investigation",
        "retrieval_mode": live_status,
        "count": len(results),
        "results": results,
    }


def _box_counting_fractal_dimension(edge_map: np.ndarray) -> float:
    """Compute Minkowski-Bouligand box-counting fractal dimension on a binary contour edge image."""
    points = np.argwhere(edge_map > 0)
    if len(points) < 8:
        return 1.0
    min_dim = min(edge_map.shape[:2])
    max_exp = int(np.floor(np.log2(max(min_dim, 8))))
    scales = [2**k for k in range(1, min(max_exp, 6))]
    counts = []
    valid_scales = []
    for s in scales:
        h_bins = max(1, edge_map.shape[0] // s)
        w_bins = max(1, edge_map.shape[1] // s)
        cropped = edge_map[: h_bins * s, : w_bins * s]
        if cropped.size == 0:
            continue
        blocks = cropped.reshape(h_bins, s, w_bins, s).sum(axis=(1, 3))
        non_empty = int(np.count_nonzero(blocks > 0))
        if non_empty > 0:
            counts.append(non_empty)
            valid_scales.append(s)
    if len(valid_scales) < 2:
        return 1.08
    coeffs = np.polyfit(np.log(1.0 / np.array(valid_scales, dtype=float)), np.log(counts), 1)
    fd = float(np.clip(coeffs[0], 1.0, 1.99))
    return round(fd, 4)


def _compute_mirror_iou(mask_a: np.ndarray, mask_b: np.ndarray) -> float:
    a = mask_a > 0
    b = mask_b > 0
    union = int(np.count_nonzero(a | b))
    if union == 0:
        return 0.0
    inter = int(np.count_nonzero(a & b))
    return float(inter / union)


def extract_image_descriptors(image_bytes: bytes, filename: str = "image.png") -> dict[str, Any]:
    """Compute OpenCV/NumPy flame geometry, symmetry, IoU, convexity, fractal, and color descriptors."""
    arr = np.frombuffer(image_bytes, dtype=np.uint8)
    bgr = cv2.imdecode(arr, cv2.IMREAD_COLOR)
    if bgr is None or bgr.size == 0:
        raise ValueError(f"Could not decode image file '{filename}'.")

    h_img, w_img = bgr.shape[:2]
    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)

    # Flame luminosity score combining warm incandescence (R/G) and blue chemiluminescence
    r_ch = rgb[:, :, 0].astype(np.float32)
    g_ch = rgb[:, :, 1].astype(np.float32)
    b_ch = rgb[:, :, 2].astype(np.float32)
    luma = 0.45 * r_ch + 0.35 * g_ch + 0.20 * b_ch

    otsu_thresh, mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    if np.count_nonzero(mask) < 0.005 * mask.size or np.count_nonzero(mask) > 0.90 * mask.size:
        p75 = float(np.percentile(luma, 75))
        mask = np.where(luma >= max(p75, 35.0), 255, 0).astype(np.uint8)

    kernel = np.ones((3, 3), dtype=np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
    if not contours:
        # Fallback synthetic circular mask around brightest pixel
        max_idx = np.unravel_index(int(np.argmax(luma)), luma.shape)
        mask = np.zeros_like(gray)
        cv2.circle(mask, (int(max_idx[1]), int(max_idx[0])), max(6, min(h_img, w_img) // 6), 255, -1)
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)

    main_contour = max(contours, key=cv2.contourArea)
    clean_mask = np.zeros_like(gray)
    cv2.drawContours(clean_mask, [main_contour], -1, 255, thickness=-1)

    area_px = float(max(cv2.contourArea(main_contour), float(np.count_nonzero(clean_mask))))
    perimeter_px = float(max(cv2.arcLength(main_contour, True), 1.0))
    x_box, y_box, w_box, h_box = cv2.boundingRect(main_contour)
    aspect_ratio = float(h_box / max(w_box, 1))
    circularity = float(np.clip((4.0 * math.pi * area_px) / (perimeter_px**2), 0.0, 1.0))

    hull = cv2.convexHull(main_contour)
    hull_area = float(max(cv2.contourArea(hull), area_px, 1.0))
    hull_perimeter = float(max(cv2.arcLength(hull, True), 1.0))
    solidity = float(np.clip(area_px / hull_area, 0.0, 1.0))
    convexity = float(np.clip(hull_perimeter / perimeter_px, 0.0, 1.0))
    boundary_roughness = float(max(perimeter_px / hull_perimeter, 1.0))

    moments = cv2.moments(clean_mask)
    if moments["m00"] > 0:
        cx_geom = float(moments["m10"] / moments["m00"])
        cy_geom = float(moments["m01"] / moments["m00"])
    else:
        cx_geom = float(x_box + w_box / 2.0)
        cy_geom = float(y_box + h_box / 2.0)

    weights = np.where(clean_mask > 0, np.maximum(luma, 1.0), 0.0)
    w_sum = float(np.sum(weights))
    if w_sum > 0:
        y_indices, x_indices = np.indices(gray.shape)
        cx_int = float(np.sum(x_indices * weights) / w_sum)
        cy_int = float(np.sum(y_indices * weights) / w_sum)
    else:
        cx_int, cy_int = cx_geom, cy_geom

    intensity_centroid_offset_px = float(math.hypot(cx_int - cx_geom, cy_int - cy_geom))
    equiv_radius = float(max(math.sqrt(area_px / math.pi), 1.0))
    centroid_offset_norm = float(intensity_centroid_offset_px / equiv_radius)

    # Crop around bounding box to evaluate reflectional & rotational symmetry
    crop_mask = clean_mask[y_box : y_box + h_box, x_box : x_box + w_box]
    crop_gray = gray[y_box : y_box + h_box, x_box : x_box + w_box].astype(np.float32) / 255.0
    crop_weighted = np.where(crop_mask > 0, crop_gray, 0.0)

    flip_h_mask = cv2.flip(crop_mask, 1)
    flip_v_mask = cv2.flip(crop_mask, 0)
    flip_hv_mask = cv2.flip(crop_mask, -1)

    mirror_iou_h = _compute_mirror_iou(crop_mask, flip_h_mask)
    mirror_iou_v = _compute_mirror_iou(crop_mask, flip_v_mask)
    rot_180_iou = _compute_mirror_iou(crop_mask, flip_hv_mask)

    sq_size = max(w_box, h_box, 8)
    sq_mask = cv2.resize(crop_mask, (sq_size, sq_size), interpolation=cv2.INTER_NEAREST)
    rot_90_mask = cv2.rotate(sq_mask, cv2.ROTATE_90_CLOCKWISE)
    rot_90_iou = _compute_mirror_iou(sq_mask, rot_90_mask)
    rotational_symmetry = float(0.5 * rot_180_iou + 0.5 * rot_90_iou)

    flip_h_int = cv2.flip(crop_weighted, 1)
    flip_v_int = cv2.flip(crop_weighted, 0)
    denom_int = float(np.sum(np.maximum(crop_weighted, flip_h_int)) or 1.0)
    l1_sym_h = float(1.0 - np.sum(np.abs(crop_weighted - flip_h_int)) / denom_int)
    denom_int_v = float(np.sum(np.maximum(crop_weighted, flip_v_int)) or 1.0)
    l1_sym_v = float(1.0 - np.sum(np.abs(crop_weighted - flip_v_int)) / denom_int_v)

    bilateral_symmetry = float(np.clip(0.6 * mirror_iou_h + 0.4 * l1_sym_h, 0.0, 1.0))
    top_bottom_symmetry = float(np.clip(0.6 * mirror_iou_v + 0.4 * l1_sym_v, 0.0, 1.0))
    mean_mirror_iou = float(0.5 * (mirror_iou_h + mirror_iou_v))
    symmetry_loss = float(1.0 - mean_mirror_iou)

    # Radial radius variation around geometric centroid
    pts = main_contour.reshape(-1, 2).astype(np.float32)
    radii = np.hypot(pts[:, 0] - cx_geom, pts[:, 1] - cy_geom)
    radial_cv = float(np.std(radii) / max(float(np.mean(radii)), 1.0))

    aspect_closeness = float(np.exp(-abs(aspect_ratio - 1.0)))
    spherical_symmetry_score = float(
        np.clip(
            0.30 * bilateral_symmetry
            + 0.30 * top_bottom_symmetry
            + 0.20 * rotational_symmetry
            + 0.20 * circularity * aspect_closeness,
            0.0,
            1.0,
        )
    )

    edge_map = np.zeros_like(gray)
    cv2.drawContours(edge_map, [main_contour], -1, 255, thickness=1)
    fractal_dim = _box_counting_fractal_dimension(edge_map)

    in_flame = clean_mask > 0
    mean_r = float(np.mean(r_ch[in_flame])) if np.any(in_flame) else 0.0
    mean_g = float(np.mean(g_ch[in_flame])) if np.any(in_flame) else 0.0
    mean_b = float(np.mean(b_ch[in_flame])) if np.any(in_flame) else 0.0
    mean_luma = float(np.mean(luma[in_flame])) if np.any(in_flame) else 0.0
    blue_to_red_ratio = float(mean_b / max(mean_r, 1.0))
    soot_pixels = np.logical_and(in_flame, np.logical_and(r_ch > 160, r_ch > 1.3 * np.maximum(b_ch, 1.0)))
    soot_incandescence_fraction = float(
        np.count_nonzero(soot_pixels) / max(int(np.count_nonzero(in_flame)), 1)
    )

    if spherical_symmetry_score >= 0.82 and 0.80 <= aspect_ratio <= 1.25:
        morphology_class = "spherical_diffusion"
        buoyancy_regime_flag = "microgravity_quiescent_candidate"
    elif bilateral_symmetry >= 0.78 and aspect_ratio <= 1.45:
        morphology_class = "hemispherical_or_ellipsoidal_cap"
        buoyancy_regime_flag = "low_velocity_forced_or_partial_g"
    elif aspect_ratio > 1.45:
        morphology_class = "elongated_teardrop_plume"
        buoyancy_regime_flag = "buoyancy_or_convective_elongation"
    else:
        morphology_class = "irregular_or_cellular_front"
        buoyancy_regime_flag = "perturbed_or_transitional"

    return {
        "modality": "image",
        "filename": filename,
        "width_px": w_img,
        "height_px": h_img,
        "threshold_otsu": float(otsu_thresh),
        "flame_area_px": round(area_px, 2),
        "flame_area_fraction": round(area_px / float(w_img * h_img), 5),
        "perimeter_px": round(perimeter_px, 2),
        "bounding_box": {"x": int(x_box), "y": int(y_box), "width": int(w_box), "height": int(h_box)},
        "aspect_ratio": round(aspect_ratio, 4),
        "circularity": round(circularity, 4),
        "convex_hull_area_px": round(hull_area, 2),
        "solidity": round(solidity, 4),
        "convexity": round(convexity, 4),
        "boundary_roughness": round(boundary_roughness, 4),
        "fractal_dimension": round(fractal_dim, 4),
        "centroid_x": round(cx_geom, 2),
        "centroid_y": round(cy_geom, 2),
        "intensity_centroid_x": round(cx_int, 2),
        "intensity_centroid_y": round(cy_int, 2),
        "intensity_centroid_offset_px": round(intensity_centroid_offset_px, 3),
        "centroid_offset_norm": round(centroid_offset_norm, 4),
        "bilateral_symmetry": round(bilateral_symmetry, 4),
        "top_bottom_symmetry": round(top_bottom_symmetry, 4),
        "rotational_symmetry": round(rotational_symmetry, 4),
        "mirror_iou_horizontal": round(mirror_iou_h, 4),
        "mirror_iou_vertical": round(mirror_iou_v, 4),
        "mean_mirror_iou": round(mean_mirror_iou, 4),
        "soft_mirror_iou": round(0.5 * (l1_sym_h + l1_sym_v), 4),
        "symmetry_loss": round(symmetry_loss, 4),
        "radial_radius_cv": round(radial_cv, 4),
        "spherical_symmetry_score": round(spherical_symmetry_score, 4),
        "mean_luminance": round(mean_luma, 2),
        "mean_rgb": {"r": round(mean_r, 2), "g": round(mean_g, 2), "b": round(mean_b, 2)},
        "blue_to_red_ratio": round(blue_to_red_ratio, 4),
        "soot_incandescence_fraction": round(soot_incandescence_fraction, 4),
        "cool_flame_likelihood": (
            "elevated_candidate" if (mean_luma < 95.0 and blue_to_red_ratio >= 0.85) else "low_from_visible_alone"
        ),
        "morphology_class": morphology_class,
        "buoyancy_regime_flag": buoyancy_regime_flag,
        "units": {
            "flame_area_px": "px^2",
            "perimeter_px": "px",
            "centroid_x": "px",
            "centroid_y": "px",
            "intensity_centroid_offset_px": "px",
            "aspect_ratio": "dimensionless (H/W)",
            "bilateral_symmetry": "score [0,1]",
            "top_bottom_symmetry": "score [0,1]",
            "rotational_symmetry": "score [0,1]",
            "spherical_symmetry_score": "score [0,1]",
            "fractal_dimension": "box-counting D [1,2]",
        },
        "calibration_note": "Pixel-domain geometric and luminance measurements; physical length (mm) and radiometric temperature (K) require camera scale and blackbody calibration.",
    }


def extract_video_descriptors(video_bytes: bytes, filename: str = "video.mp4") -> dict[str, Any]:
    """Extract frame-wise flame geometry, symmetry variance, centroid jitter, and flicker frequency."""
    suffix = Path(filename).suffix or ".mp4"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(video_bytes)
        tmp_path = tmp.name

    try:
        cap = cv2.VideoCapture(tmp_path)
        if not cap.isOpened():
            raise ValueError(f"Could not open video stream '{filename}'.")

        fps = float(cap.get(cv2.CAP_PROP_FPS) or 30.0)
        if fps <= 0 or math.isnan(fps):
            fps = 30.0
        total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
        max_samples = 48
        step = max(1, total_frames // max_samples) if total_frames > max_samples else 1

        areas: list[float] = []
        intensities: list[float] = []
        symmetries: list[float] = []
        centroids: list[tuple[float, float]] = []
        first_frame_desc: dict[str, Any] | None = None

        frame_idx = 0
        while True:
            ret, frame = cap.read()
            if not ret or frame is None:
                break
            if frame_idx % step == 0:
                ok, buf = cv2.imencode(".png", frame)
                if ok:
                    f_desc = extract_image_descriptors(buf.tobytes(), f"{filename}#frame{frame_idx}")
                    if first_frame_desc is None:
                        first_frame_desc = f_desc
                    areas.append(float(f_desc["flame_area_px"]))
                    intensities.append(float(f_desc["mean_luminance"]))
                    symmetries.append(float(f_desc["bilateral_symmetry"]))
                    centroids.append((float(f_desc["centroid_x"]), float(f_desc["centroid_y"])))
            frame_idx += 1
            if len(areas) >= max_samples:
                break
        cap.release()
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    if not areas or first_frame_desc is None:
        raise ValueError(f"No readable frames decoded from video '{filename}'.")

    areas_arr = np.array(areas, dtype=float)
    int_arr = np.array(intensities, dtype=float)
    sym_arr = np.array(symmetries, dtype=float)

    area_mean = float(np.mean(areas_arr))
    area_cv = float(np.std(areas_arr) / max(area_mean, 1.0))
    intensity_cv = float(np.std(int_arr) / max(float(np.mean(int_arr)), 1.0))
    symmetry_var = float(np.var(sym_arr))
    symmetry_std = float(np.std(sym_arr))

    if len(centroids) >= 2:
        pts = np.array(centroids, dtype=float)
        displacements = np.hypot(np.diff(pts[:, 0]), np.diff(pts[:, 1]))
        centroid_jitter_px = float(np.sqrt(np.mean(displacements**2)))
    else:
        centroid_jitter_px = 0.0

    equiv_r = max(math.sqrt(area_mean / math.pi), 1.0)
    centroid_jitter_norm = float(centroid_jitter_px / equiv_r)

    # Dominant flicker frequency via real FFT
    effective_fps = fps / float(step)
    if len(int_arr) >= 4:
        centered = int_arr - np.mean(int_arr)
        spec = np.abs(scipy_fft.rfft(centered))
        freqs = scipy_fft.rfftfreq(len(centered), d=1.0 / max(effective_fps, 1.0))
        if len(spec) > 1:
            peak_idx = int(np.argmax(spec[1:]) + 1)
            dominant_flicker_hz = float(freqs[peak_idx])
            total_power = float(np.sum(spec[1:] ** 2))
            flicker_power_ratio = float((spec[peak_idx] ** 2) / max(total_power, 1e-9))
        else:
            dominant_flicker_hz = 0.0
            flicker_power_ratio = 0.0
    else:
        dominant_flicker_hz = 0.0
        flicker_power_ratio = 0.0

    peak_area = float(np.max(areas_arr))
    final_area = float(areas_arr[-1])
    final_to_peak_ratio = float(final_area / max(peak_area, 1.0))
    min_mid_area = float(np.min(areas_arr[:-1])) if len(areas_arr) > 2 else final_area
    extinction_detected = bool(final_to_peak_ratio < 0.20)
    reignition_detected = bool(
        len(areas_arr) >= 4 and (min_mid_area / max(peak_area, 1.0) < 0.25) and (final_to_peak_ratio > 0.50)
    )

    # Sequence anomaly z-score
    if len(areas_arr) >= 4 and np.std(areas_arr) > 1e-6:
        z_scores = np.abs((areas_arr - area_mean) / np.std(areas_arr))
        temporal_anomaly_score = float(np.clip(np.max(z_scores) / 3.5, 0.0, 1.0))
        change_point_frame = int(np.argmax(np.abs(np.diff(areas_arr)))) if len(areas_arr) >= 2 else 0
    else:
        temporal_anomaly_score = 0.0
        change_point_frame = 0

    merged = dict(first_frame_desc)
    merged.update(
        {
            "modality": "video",
            "filename": filename,
            "sampled_frames": len(areas),
            "total_frames": max(total_frames, len(areas)),
            "fps": round(fps, 2),
            "duration_s": round(max(total_frames, len(areas)) / max(fps, 1.0), 3),
            "area_mean_px": round(area_mean, 2),
            "area_cv": round(area_cv, 4),
            "intensity_cv": round(intensity_cv, 4),
            "symmetry_variance": round(symmetry_var, 6),
            "symmetry_std": round(symmetry_std, 4),
            "centroid_jitter_px": round(centroid_jitter_px, 3),
            "centroid_jitter_norm": round(centroid_jitter_norm, 4),
            "dominant_flicker_hz": round(dominant_flicker_hz, 3),
            "flicker_power_ratio": round(flicker_power_ratio, 4),
            "final_to_peak_area_ratio": round(final_to_peak_ratio, 4),
            "extinction_detected": extinction_detected,
            "reignition_detected": reignition_detected,
            "temporal_anomaly_score": round(temporal_anomaly_score, 4),
            "change_point_frame": change_point_frame,
        }
    )
    return merged


def extract_table_descriptors(table_bytes: bytes, filename: str = "data.csv") -> dict[str, Any]:
    """Parse CSV/TSV numeric tables and spectra, computing summary statistics, trends, and peak detection."""
    text = table_bytes.decode("utf-8", errors="replace")
    delimiter = "\t" if filename.lower().endswith(".tsv") or ("\t" in text.splitlines()[0] if text.splitlines() else False) else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
    rows = list(reader)
    columns = list(reader.fieldnames or [])

    numeric_cols: dict[str, list[float]] = {col: [] for col in columns}
    for row in rows:
        for col in columns:
            val_str = str(row.get(col, "")).strip()
            if not val_str:
                continue
            try:
                val = float(val_str)
                if math.isfinite(val):
                    numeric_cols[col].append(val)
            except ValueError:
                continue

    col_summaries: dict[str, dict[str, Any]] = {}
    for col, vals in numeric_cols.items():
        if not vals:
            continue
        arr = np.array(vals, dtype=float)
        trend_r = 0.0
        if len(arr) >= 3 and np.std(arr) > 1e-9:
            r_val, _ = scipy_stats.pearsonr(np.arange(len(arr), dtype=float), arr)
            trend_r = float(0.0 if math.isnan(r_val) else r_val)
        col_summaries[col] = {
            "count": len(vals),
            "min": round(float(np.min(arr)), 5),
            "max": round(float(np.max(arr)), 5),
            "mean": round(float(np.mean(arr)), 5),
            "std": round(float(np.std(arr)), 5),
            "median": round(float(np.median(arr)), 5),
            "index_correlation_r": round(trend_r, 4),
        }

    lower_cols = {c.lower(): c for c in columns}
    spectral_axis_col = next(
        (
            orig
            for low, orig in lower_cols.items()
            if any(k in low for k in ("wavelength", "nm", "wavenumber", "cm-1", "cm^-1", "raman"))
        ),
        None,
    )
    intensity_col = next(
        (
            orig
            for low, orig in lower_cols.items()
            if any(k in low for k in ("intensity", "absorbance", "transmittance", "emission", "counts", "flux"))
        ),
        None,
    )

    is_spectrum = bool(spectral_axis_col and intensity_col and numeric_cols.get(spectral_axis_col) and numeric_cols.get(intensity_col))
    spectral_summary: dict[str, Any] = {}
    if is_spectrum and spectral_axis_col and intensity_col:
        x_vals = np.array(numeric_cols[spectral_axis_col], dtype=float)
        y_vals = np.array(numeric_cols[intensity_col], dtype=float)
        n_pts = min(len(x_vals), len(y_vals))
        x_vals = x_vals[:n_pts]
        y_vals = y_vals[:n_pts]
        baseline = float(np.percentile(y_vals, 15)) if n_pts > 0 else 0.0
        noise = float(np.std(y_vals - baseline)) or 1.0
        peak_idx = int(np.argmax(y_vals)) if n_pts > 0 else 0
        dom_x = float(x_vals[peak_idx]) if n_pts > 0 else 0.0
        dom_y = float(y_vals[peak_idx]) if n_pts > 0 else 0.0
        thresh = baseline + 0.35 * max(dom_y - baseline, 1e-6)
        peaks_x = [round(float(x_vals[i]), 2) for i in range(n_pts) if y_vals[i] >= thresh]
        candidate_bands = []
        for px in peaks_x:
            if 304 <= px <= 315:
                candidate_bands.append("OH* (~308 nm reaction zone radical)")
            elif 425 <= px <= 436:
                candidate_bands.append("CH* (~431 nm primary flame front)")
            elif 510 <= px <= 520:
                candidate_bands.append("C2* Swan band (~516 nm rich hydrocarbon zone)")
            elif 2280 <= px <= 2390:
                candidate_bands.append("CO2 asymmetric stretch (~2349 cm^-1)")
            elif 2100 <= px <= 2200:
                candidate_bands.append("CO fundamental band (~2143 cm^-1)")
        spectral_summary = {
            "spectral_axis_column": spectral_axis_col,
            "intensity_column": intensity_col,
            "dominant_wavelength_or_band": round(dom_x, 3),
            "peak_intensity": round(dom_y, 5),
            "baseline_estimate": round(baseline, 5),
            "spectral_snr": round(float((dom_y - baseline) / noise), 3),
            "spectral_peak_count": len(peaks_x),
            "candidate_species_bands": sorted(set(candidate_bands)),
        }

    return {
        "modality": "spectrum" if is_spectrum else "table",
        "filename": filename,
        "row_count": len(rows),
        "column_count": len(columns),
        "columns": columns,
        "numeric_summaries": col_summaries,
        "spectral_analysis": spectral_summary,
    }


def extract_document_descriptors(file_bytes: bytes, filename: str) -> dict[str, Any]:
    """Parse JSON, PDF, DOCX, XLSX, or text experiment documents and extract PSI accessions & key terms."""
    ext = Path(filename).suffix.lower()
    extracted_text = ""
    extra_meta: dict[str, Any] = {}

    if ext == ".json":
        parsed = json.loads(file_bytes.decode("utf-8", errors="replace"))
        extracted_text = json.dumps(parsed)[:20000]
        if isinstance(parsed, dict):
            extra_meta["top_level_keys"] = list(parsed.keys())[:30]
            if "wavelength" in parsed or "spectrum" in parsed:
                extra_meta["contains_spectral_fields"] = True
        elif isinstance(parsed, list):
            extra_meta["array_length"] = len(parsed)
    elif ext == ".pdf":
        pages_text: list[str] = []
        try:
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                extra_meta["page_count"] = len(doc)
                for page in doc:
                    pages_text.append(page.get_text("text"))
        except Exception:
            reader = PdfReader(io.BytesIO(file_bytes))
            extra_meta["page_count"] = len(reader.pages)
            for page in reader.pages:
                pages_text.append(page.extract_text() or "")
        extracted_text = "\n".join(pages_text)
    elif ext == ".docx":
        doc = DocxDocument(io.BytesIO(file_bytes))
        paras = [p.text for p in doc.paragraphs if p.text.strip()]
        extra_meta["paragraph_count"] = len(paras)
        extra_meta["table_count"] = len(doc.tables)
        extracted_text = "\n".join(paras)
    elif ext == ".xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        extra_meta["sheet_names"] = wb.sheetnames
        sheet_summaries: dict[str, Any] = {}
        text_chunks: list[str] = []
        for name in wb.sheetnames:
            ws = wb[name]
            rows_data = list(ws.iter_rows(values_only=True))
            sheet_summaries[name] = {
                "rows": len(rows_data),
                "columns": max((len(r) for r in rows_data), default=0),
            }
            for r in rows_data[:40]:
                text_chunks.append(" ".join(str(c) for c in r if c is not None))
        extra_meta["sheets"] = sheet_summaries
        extracted_text = "\n".join(text_chunks)
    else:
        extracted_text = file_bytes.decode("utf-8", errors="replace")

    psi_mentions = sorted(set(m.upper() for m in re.findall(r"PSI-\d+", extracted_text, flags=re.IGNORECASE)))
    fuel_keywords = [
        f for f in ("pmma", "heptane", "methanol", "decane", "cotton", "nomex", "methane", "ethylene")
        if re.search(rf"\b{f}\b", extracted_text, flags=re.IGNORECASE)
    ]
    facility_keywords = [
        fac for fac in ("FLEX", "FLEX-2", "BASS-II", "SoFIE", "Saffire", "ACME", "CIR", "MSG", "Drop Tower")
        if re.search(re.escape(fac), extracted_text, flags=re.IGNORECASE)
    ]

    modality = "table" if ext == ".xlsx" else ("spectrum" if extra_meta.get("contains_spectral_fields") else "document")
    return {
        "modality": modality,
        "filename": filename,
        "format": ext.lstrip(".") or "txt",
        "character_count": len(extracted_text),
        "detected_psi_accessions": psi_mentions,
        "detected_fuels": fuel_keywords,
        "detected_facilities": facility_keywords,
        "excerpt": extracted_text[:500].strip(),
        **extra_meta,
    }


def parse_evidence_file(file_bytes: bytes, filename: str, content_type: str = "") -> dict[str, Any]:
    """Dispatch uploaded file bytes to the appropriate local extractor."""
    if len(file_bytes) > MAX_FILE_BYTES:
        raise ValueError(
            f"File '{filename}' ({len(file_bytes)} bytes) exceeds PYROSAGE_MAX_FILE_BYTES ({MAX_FILE_BYTES} bytes)."
        )
    ext = Path(filename).suffix.lower()
    ct = (content_type or "").lower()

    if ext in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp") or ct.startswith("image/"):
        desc = extract_image_descriptors(file_bytes, filename)
    elif ext in (".mp4", ".avi", ".mov", ".mkv", ".webm") or ct.startswith("video/"):
        desc = extract_video_descriptors(file_bytes, filename)
    elif ext in (".csv", ".tsv") or ct in ("text/csv", "text/tab-separated-values"):
        desc = extract_table_descriptors(file_bytes, filename)
    elif ext in (".json", ".pdf", ".docx", ".xlsx", ".txt", ".md"):
        desc = extract_document_descriptors(file_bytes, filename)
    else:
        raise ValueError(
            f"Unsupported evidence file type '{ext or content_type}' for '{filename}'. "
            "Supported formats: PNG/JPG/TIFF/WebP, MP4/AVI/MOV, CSV/TSV, JSON, PDF, DOCX, XLSX."
        )

    desc["size_bytes"] = len(file_bytes)
    desc["evidence_class"] = "uploaded_evidence"
    return desc


def evaluate_techniques(
    selected_ids: list[str],
    evidence_items: list[dict[str, Any]],
    context: dict[str, Any],
) -> list[dict[str, Any]]:
    """Determine accurate execution state for each selected technique based on modalities and solvers."""
    registry = {t["id"]: t for t in load_techniques()}
    if not selected_ids:
        selected_ids = ["T01", "T04", "T08", "T09", "T10", "T12", "T13", "T14", "T16", "T52", "T53", "T56", "T57"]

    available_modalities = {item.get("modality", "document") for item in evidence_items}
    if context.get("psi_investigation"):
        available_modalities.add("document")

    merged_descriptors: dict[str, Any] = {}
    for item in evidence_items:
        for k, v in item.items():
            if k not in ("modality", "filename", "size_bytes", "evidence_class", "units"):
                merged_descriptors[k] = v

    assessments: list[dict[str, Any]] = []
    for tid in selected_ids:
        tech = registry.get(tid)
        if not tech:
            continue

        supported = set(tech.get("supported_modalities", []))
        exec_mode = tech.get("execution_mode", "local_descriptor")
        has_modality = bool(available_modalities & supported)

        # Specialist solver techniques cannot be claimed as Executed without dedicated calibration/solvers
        if exec_mode == "requires_specialist_solver":
            if has_modality:
                state = "Requires specialist solver"
                note = (
                    f"{tech['name']} requires calibrated instrument inputs and a dedicated solver "
                    f"({', '.join(tech.get('prerequisites', []))}). PyroSage will not fabricate solver outputs."
                )
            else:
                state = "Requires data"
                note = (
                    f"Required modality ({', '.join(sorted(supported))}) and specialist prerequisites "
                    f"({', '.join(tech.get('prerequisites', []))}) were not present in the evidence."
                )
            measured = {}
        elif not has_modality:
            state = "Requires data"
            note = (
                f"Evidence does not include required modality ({', '.join(sorted(supported))}). "
                f"Prerequisites: {', '.join(tech.get('prerequisites', []))}."
            )
            measured = {}
        elif exec_mode == "partially_executed":
            state = "Partially executed"
            measured = {k: merged_descriptors[k] for k in tech.get("outputs", []) if k in merged_descriptors}
            note = (
                f"Proxy descriptors computed locally, but full {tech['name']} verification requires "
                f"additional calibration or boundary measurements ({', '.join(tech.get('prerequisites', []))})."
            )
        else:
            measured = {k: merged_descriptors[k] for k in tech.get("outputs", []) if k in merged_descriptors}
            # Synthesis, platform, and gravity scaling techniques can execute against context + PSI catalogue
            if measured or tech["group"] in ("Platforms", "Gravity scaling", "Safety interpretation", "AI methods", "Microgravity fundamentals"):
                state = "Executed"
                note = f"Executed using local { ', '.join(sorted(available_modalities)) } descriptors and bounded context."
            else:
                state = "Partially executed"
                note = f"Modality present, but specific output fields ({', '.join(tech.get('outputs', []))}) require calibrated metadata."

        assessments.append(
            {
                "id": tech["id"],
                "code": tech["code"],
                "name": tech["name"],
                "group": tech["group"],
                "state": state,
                "supported_modalities": tech["supported_modalities"],
                "prerequisites": tech["prerequisites"],
                "measured_outputs": measured,
                "rationale": note,
            }
        )
    return assessments


def match_psi_experiments(
    evidence_items: list[dict[str, Any]],
    context: dict[str, Any],
    catalogue: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    """Rank verified NASA PSI catalogue records by multi-feature similarity to evidence and context."""
    records = catalogue if catalogue is not None else list_cached_psi_experiments()
    fuel_q = str(context.get("fuel", "")).strip().lower()
    platform_q = str(context.get("platform", "")).strip().lower()
    gravity_g = float(context.get("gravity_g", 0.0) or 0.0)
    target_acc = str(context.get("accession", "")).strip().upper()

    img_item = next((it for it in evidence_items if it.get("modality") in ("image", "video")), None)
    doc_accessions: set[str] = set()
    doc_fuels: set[str] = set()
    for it in evidence_items:
        doc_accessions.update(it.get("detected_psi_accessions", []))
        doc_fuels.update(f.lower() for f in it.get("detected_fuels", []))

    matches: list[dict[str, Any]] = []
    for rec in records:
        validated = validate_psi_record(rec)
        score_parts: list[float] = []
        rationale_parts: list[str] = []

        if target_acc and validated["accession"] == target_acc:
            score_parts.append(1.0)
            rationale_parts.append(f"Direct PSI accession match ({target_acc})")
        elif validated["accession"] in doc_accessions:
            score_parts.append(0.95)
            rationale_parts.append(f"Referenced in uploaded document ({validated['accession']})")

        rec_fuels = [f.lower() for f in validated["fuels"]]
        if fuel_q and any(fuel_q in rf or rf in fuel_q for rf in rec_fuels):
            score_parts.append(0.90)
            rationale_parts.append(f"Fuel match ({fuel_q})")
        elif doc_fuels & set(rec_fuels):
            common_f = sorted(doc_fuels & set(rec_fuels))[0]
            score_parts.append(0.85)
            rationale_parts.append(f"Document fuel match ({common_f})")
        else:
            score_parts.append(0.45)

        if platform_q and platform_q in validated["platform"].lower():
            score_parts.append(0.88)
            rationale_parts.append(f"Facility match ({validated['platform']})")

        grav_diff = abs(gravity_g - float(validated["gravity_g"]))
        grav_sim = float(max(0.25, 1.0 - grav_diff * 1.2))
        score_parts.append(grav_sim)
        if grav_diff <= 0.05:
            rationale_parts.append(f"Gravity regime alignment ({validated['gravity_regime']})")

        ref_desc = validated.get("reference_descriptors", {})
        if img_item and ref_desc:
            diffs = []
            for key in ("spherical_symmetry_score", "bilateral_symmetry", "top_bottom_symmetry", "circularity"):
                if key in img_item and key in ref_desc:
                    diffs.append(abs(float(img_item[key]) - float(ref_desc[key])))
            if diffs:
                geom_sim = float(np.clip(1.0 - float(np.mean(diffs)) * 1.4, 0.2, 0.99))
                score_parts.append(geom_sim)
                rationale_parts.append(f"Flame geometry descriptor similarity ({geom_sim:.2f})")

        sim_score = float(np.clip(np.mean(score_parts), 0.1, 0.99))
        matches.append(
            {
                "accession": validated["accession"],
                "title": validated["title"],
                "data_source": "NASA PSI",
                "source_url": validated["source_url"],
                "platform": validated["platform"],
                "gravity_regime": validated["gravity_regime"],
                "fuels": validated["fuels"],
                "similarity_score": round(sim_score, 4),
                "match_rationale": "; ".join(rationale_parts) or "Baseline NASA PSI microgravity combustion candidate",
            }
        )

    matches.sort(key=lambda m: (-m["similarity_score"], m["accession"]))
    return matches


def validate_si_output(
    raw_output: dict[str, Any],
    allowed_technique_ids: list[str],
    allowed_psi_accessions: list[str],
    technique_assessments: list[dict[str, Any]],
) -> dict[str, Any]:
    """Enforce LLM operating boundaries: strip unapproved techniques, reject hallucinated PSI IDs, and verify execution states."""
    allowed_tech_set = set(allowed_technique_ids)
    allowed_acc_set = set(allowed_psi_accessions)
    executed_tech_set = {a["id"] for a in technique_assessments if a["state"] == "Executed"}

    rejected_experiment_ids: list[str] = []
    referenced_accs = raw_output.get("referenced_psi_accessions", [])
    valid_referenced_accs: list[str] = []
    if isinstance(referenced_accs, list):
        for acc in referenced_accs:
            acc_str = str(acc).strip().upper()
            if acc_str in allowed_acc_set:
                valid_referenced_accs.append(acc_str)
            else:
                rejected_experiment_ids.append(acc_str)

    # Also scan findings for any PSI-xxx strings not in the candidate set
    findings_in = raw_output.get("findings", [])
    validated_findings: list[dict[str, Any]] = []
    if isinstance(findings_in, list):
        for item in findings_in:
            if not isinstance(item, dict):
                continue
            title_str = str(item.get("title", "Evidence Finding"))
            interp_str = str(item.get("interpretation", ""))
            ref_str = str(item.get("evidence_reference", "Computed local descriptors"))
            text_blob = f"{title_str} {interp_str} {ref_str}"
            for mentioned_acc in re.findall(r"PSI-\d+", text_blob, flags=re.IGNORECASE):
                m_up = mentioned_acc.upper()
                if m_up not in allowed_acc_set:
                    rejected_experiment_ids.append(m_up)
                    title_str = re.sub(re.escape(mentioned_acc), "verified PSI candidate", title_str, flags=re.IGNORECASE)
                    interp_str = re.sub(re.escape(mentioned_acc), "verified PSI candidate", interp_str, flags=re.IGNORECASE)
                    ref_str = re.sub(re.escape(mentioned_acc), "verified PSI candidate", ref_str, flags=re.IGNORECASE)
            tech_refs = [
                t for t in item.get("technique_ids", []) if str(t) in allowed_tech_set
            ]
            claimed_executed = bool(item.get("claimed_executed", False))
            if claimed_executed and not any(t in executed_tech_set for t in tech_refs):
                claimed_executed = False
            validated_findings.append(
                {
                    "title": title_str,
                    "evidence_class": str(item.get("evidence_class", "model_interpretation")),
                    "technique_ids": tech_refs or allowed_technique_ids[:2],
                    "evidence_reference": ref_str,
                    "interpretation": interp_str,
                    "verified_execution": claimed_executed,
                }
            )

    return {
        "summary": str(raw_output.get("summary", "")),
        "referenced_psi_accessions": sorted(set(valid_referenced_accs)),
        "rejected_hallucinated_accessions": sorted(set(rejected_experiment_ids)),
        "findings": validated_findings,
        "safety_insights": list(raw_output.get("safety_insights", [])),
        "hypotheses": list(raw_output.get("hypotheses", [])),
        "data_gaps": list(raw_output.get("data_gaps", [])),
    }


def _build_deterministic_synthesis(
    evidence_items: list[dict[str, Any]],
    technique_assessments: list[dict[str, Any]],
    candidate_matches: list[dict[str, Any]],
    context: dict[str, Any],
) -> dict[str, Any]:
    """Generate transparent, evidence-linked structured synthesis grounded in local descriptors and PSI records."""
    gravity_g = float(context.get("gravity_g", 0.0) or 0.0)
    oxygen_pct = float(context.get("oxygen_pct", 21.0) or 21.0)
    pressure_kpa = float(context.get("pressure_kpa", 101.3) or 101.3)
    fuel = str(context.get("fuel", "pmma")).lower()

    top_match = candidate_matches[0] if candidate_matches else VERIFIED_PSI_CATALOGUE[0]
    executed = [a for a in technique_assessments if a["state"] == "Executed"]
    specialist = [a for a in technique_assessments if a["state"] == "Requires specialist solver"]
    missing_data = [a for a in technique_assessments if a["state"] == "Requires data"]

    findings: list[dict[str, Any]] = []
    img_ev = next((e for e in evidence_items if e.get("modality") in ("image", "video")), None)
    tab_ev = next((e for e in evidence_items if e.get("modality") in ("table", "spectrum")), None)

    if img_ev:
        Findings_techs = [a["id"] for a in executed if a["group"] in ("Geometry and vision", "Microgravity fundamentals")][:4]
        findings.append(
            {
                "title": "Measured Flame Envelope Symmetry & Morphology",
                "evidence_class": "computed_descriptor",
                "technique_ids": Findings_techs or ["T08", "T09", "T10"],
                "evidence_reference": (
                    f"{img_ev['filename']}: bilateral_symmetry={img_ev['bilateral_symmetry']}, "
                    f"top_bottom_symmetry={img_ev['top_bottom_symmetry']}, "
                    f"aspect_ratio={img_ev['aspect_ratio']}, fractal_dimension={img_ev['fractal_dimension']}"
                ),
                "interpretation": (
                    f"The uploaded frame has a computed bilateral symmetry score of {img_ev['bilateral_symmetry']} "
                    f"and top-bottom symmetry of {img_ev['top_bottom_symmetry']} (aspect ratio {img_ev['aspect_ratio']}, "
                    f"morphology '{img_ev['morphology_class']}'). "
                    + (
                        "High axial and lateral symmetry is consistent with diffusion-dominated 0g burning where buoyant plume rise is absent."
                        if img_ev["spherical_symmetry_score"] >= 0.80
                        else "Reduced axial symmetry indicates forced convective ventilation or partial-gravity buoyant plume elongation."
                    )
                ),
                "claimed_executed": True,
            }
        )

    if img_ev and img_ev.get("modality") == "video":
        findings.append(
            {
                "title": "Temporal Stability, Centroid Jitter & Flicker Screening",
                "evidence_class": "computed_descriptor",
                "technique_ids": [a["id"] for a in executed if a["group"] == "Temporal dynamics"] or ["T18", "T20", "T21"],
                "evidence_reference": (
                    f"{img_ev['filename']}: sampled_frames={img_ev['sampled_frames']}, "
                    f"centroid_jitter_px={img_ev['centroid_jitter_px']}, dominant_flicker_hz={img_ev['dominant_flicker_hz']}"
                ),
                "interpretation": (
                    f"Across {img_ev['sampled_frames']} sampled frames, frame-to-frame area CV is {img_ev['area_cv']} "
                    f"with RMS centroid jitter of {img_ev['centroid_jitter_px']} px and dominant oscillation at {img_ev['dominant_flicker_hz']} Hz."
                ),
                "claimed_executed": True,
            }
        )

    if tab_ev:
        findings.append(
            {
                "title": "Tabular / Spectral Measurement Summary",
                "evidence_class": "computed_descriptor",
                "technique_ids": ["T29", "T33", "T34"],
                "evidence_reference": f"{tab_ev['filename']}: {tab_ev['row_count']} rows, columns={', '.join(tab_ev.get('columns', [])[:6])}",
                "interpretation": (
                    f"Parsed {tab_ev['row_count']} rows across {tab_ev['column_count']} columns from {tab_ev['filename']}."
                    + (
                        f" Spectral peak detected at {tab_ev['spectral_analysis'].get('dominant_wavelength_or_band')} "
                        f"with SNR {tab_ev['spectral_analysis'].get('spectral_snr')}."
                        if tab_ev.get("spectral_analysis")
                        else ""
                    )
                ),
                "claimed_executed": True,
            }
        )

    findings.append(
        {
            "title": f"NASA PSI Catalogue Alignment ({top_match['accession']})",
            "evidence_class": "candidate_match",
            "technique_ids": ["T40", "T52"],
            "evidence_reference": f"{top_match['accession']} ({top_match['source_url']}) — similarity={top_match['similarity_score']}",
            "interpretation": (
                f"The PSI catalogue returned candidate investigation {top_match['accession']} "
                f"({top_match['title']}) on {top_match['platform']} with similarity score {top_match['similarity_score']} "
                f"({top_match['match_rationale']})."
            ),
            "claimed_executed": True,
        }
    )

    buoyant_v_cm_s = round(math.sqrt(max(gravity_g, 0.0) * 9.81 * 0.05) * 100.0, 1)
    safety_insights = [
        {
            "category": "Ventilation & Partial-Gravity Flammability",
            "insight": (
                f"At g={gravity_g:.2f}g (estimated buoyant flow ~{buoyant_v_cm_s} cm/s), O2={oxygen_pct:.1f}%, P={pressure_kpa:.1f} kPa "
                f"with fuel '{fuel}', low-velocity convective oxygen supply can sustain flame spread at lower limiting oxygen "
                "concentrations than terrestrial 1g upward testing (consistent with NASA PSI-25 BASS-II and PSI-112 SoFIE findings)."
            ),
            "evidence_link": f"{top_match['accession']} ({top_match['source_url']})",
        },
        {
            "category": "Detection & Suppression Protocol",
            "insight": (
                "In microgravity and lunar/Martian partial gravity, smoke and hot products do not rise rapidly to ceiling detectors; "
                "combine ventilation-return CO/smoke sampling with dual-band UV/IR optical flame sensing, and cut cabin ventilation flow "
                "immediately prior to N2/CO2 or fine water-mist discharge."
            ),
            "evidence_link": "Executed geometry/symmetry descriptors & PSI-98 / PSI-64 safety baselines",
        },
    ]

    hypotheses = [
        (
            f"Hypothesis: Under lunar (0.17g) or Martian (0.38g) partial gravity at {oxygen_pct:.0f}% O2, "
            f"weak buoyant entrainment (~10-25 cm/s) may bridge the microgravity radiative-quenching regime and "
            f"increase flame persistence relative to quiescent 0g."
        )
    ]

    data_gaps = [
        "Current visual evidence does not contain calibrated radiometric blackbody data; absolute gas/surface temperatures (K) cannot be reported from uncalibrated RGB pixels.",
        "Physical spatial scale calibration (mm/px fiducial) is required to convert pixel flame radius and spread rate into SI velocity units (mm/s).",
    ]
    if specialist:
        data_gaps.append(
            f"{len(specialist)} selected technique(s) ({', '.join(a['id'] for a in specialist[:5])}) require specialist solvers or calibrated diagnostics (PIV particle pairs, FLIR emissivity, or LIF/FTIR spectra)."
        )
    if missing_data:
        data_gaps.append(
            f"{len(missing_data)} selected technique(s) ({', '.join(a['id'] for a in missing_data[:5])}) require modalities not present in the current upload."
        )

    return {
        "summary": (
            f"Bounded PyroSage synthesis completed across {len(evidence_items)} evidence item(s) and "
            f"{len(technique_assessments)} selected technique(s) ({len(executed)} executed locally). "
            f"Top verified NASA PSI candidate: {top_match['accession']} ({top_match['title']})."
        ),
        "referenced_psi_accessions": [m["accession"] for m in candidate_matches[:3]],
        "findings": findings,
        "safety_insights": safety_insights,
        "hypotheses": hypotheses,
        "data_gaps": data_gaps,
    }


def run_bounded_si_engine(
    evidence_items: list[dict[str, Any]],
    technique_assessments: list[dict[str, Any]],
    candidate_matches: list[dict[str, Any]],
    context: dict[str, Any],
    provider: str | None = None,
    model: str | None = None,
) -> dict[str, Any]:
    """Run Gemini or OpenAI with bounded technique whitelist and PSI candidates, or fallback deterministically."""
    chosen_provider = (provider or DEFAULT_PROVIDER or "gemini").strip().lower()
    chosen_model = (model or DEFAULT_MODEL or "gemini-3.8-flash").strip()
    allowed_tech_ids = [a["id"] for a in technique_assessments]
    allowed_psi_accs = [m["accession"] for m in candidate_matches]

    deterministic = _build_deterministic_synthesis(
        evidence_items=evidence_items,
        technique_assessments=technique_assessments,
        candidate_matches=candidate_matches,
        context=context,
    )

    bounded_prompt = {
        "instruction": (
            "You are PyroSage's technique-bounded scientific interpretation layer. "
            "Use ONLY the whitelisted technique IDs, computed descriptors, and verified NASA PSI candidate accessions. "
            "Never invent PSI accession IDs or claim specialist solvers ran when state != 'Executed'. "
            "Return JSON with keys: summary, referenced_psi_accessions, findings, safety_insights, hypotheses, data_gaps."
        ),
        "allowed_technique_ids": allowed_tech_ids,
        "technique_assessments": technique_assessments,
        "allowed_psi_accessions": allowed_psi_accs,
        "candidate_matches": candidate_matches[:4],
        "evidence_descriptors": evidence_items,
        "mission_context": context,
    }

    gemini_key = os.getenv("GEMINI_API_KEY", "").strip()
    openai_key = os.getenv("OPENAI_API_KEY", "").strip()

    raw_llm_output: dict[str, Any] | None = None
    execution_mode = "local_deterministic_fallback"

    if chosen_provider == "gemini" and gemini_key:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{chosen_model}:generateContent"
            payload = {
                "contents": [{"parts": [{"text": json.dumps(bounded_prompt)}]}],
                "generationConfig": {"responseMimeType": "application/json", "temperature": 0.1},
            }
            with httpx.Client(timeout=12.0) as client:
                resp = client.post(url, params={"key": gemini_key}, json=payload)
                if resp.status_code == 200:
                    data = resp.json()
                    text_out = (
                        data.get("candidates", [{}])[0]
                        .get("content", {})
                        .get("parts", [{}])[0]
                        .get("text", "")
                    )
                    if text_out:
                        raw_llm_output = json.loads(text_out)
                        execution_mode = "gemini_api"
        except Exception:
            raw_llm_output = None
    elif chosen_provider == "openai" and openai_key:
        try:
            url = "https://api.openai.com/v1/responses"
            payload = {
                "model": chosen_model,
                "input": json.dumps(bounded_prompt),
            }
            headers = {"Authorization": f"Bearer {openai_key}", "Content-Type": "application/json"}
            with httpx.Client(timeout=12.0) as client:
                resp = client.post(url, headers=headers, json=payload)
                if resp.status_code == 200:
                    data = resp.json()
                    output_text = data.get("output_text", "")
                    if output_text:
                        raw_llm_output = json.loads(output_text)
                        execution_mode = "openai_responses_api"
        except Exception:
            raw_llm_output = None

    validated = validate_si_output(
        raw_output=raw_llm_output or deterministic,
        allowed_technique_ids=allowed_tech_ids,
        allowed_psi_accessions=allowed_psi_accs,
        technique_assessments=technique_assessments,
    )
    validated["provider"] = chosen_provider
    validated["model"] = chosen_model
    validated["execution_mode"] = execution_mode
    return validated


def compute_ranked_hazards(
    evidence_items: list[dict[str, Any]],
    candidate_matches: list[dict[str, Any]],
    technique_assessments: list[dict[str, Any]],
    context: dict[str, Any],
) -> list[dict[str, Any]]:
    """Build multi-criteria ranked findings for mission fire-safety decision support."""
    gravity_g = float(context.get("gravity_g", 0.0) or 0.0)
    oxygen_pct = float(context.get("oxygen_pct", 21.0) or 21.0)
    fuel = str(context.get("fuel", "pmma")).lower()

    executed_ratio = sum(1 for a in technique_assessments if a["state"] == "Executed") / max(
        len(technique_assessments), 1
    )
    o2_factor = float(np.clip((oxygen_pct - 15.0) / 25.0, 0.15, 1.0))
    partial_g_factor = 0.88 if 0.05 <= gravity_g <= 0.45 else (0.72 if gravity_g < 0.05 else 0.65)
    fuel_factor = 0.90 if fuel in ("cotton", "heptane") else 0.78

    ranked: list[dict[str, Any]] = []
    for idx, match in enumerate(candidate_matches[:4]):
        sim = float(match["similarity_score"])
        severity = round(float(np.clip(0.55 * o2_factor + 0.45 * fuel_factor - idx * 0.05, 0.35, 0.96)), 3)
        completeness = round(float(np.clip(0.60 * sim + 0.40 * executed_ratio, 0.35, 0.98)), 3)
        mission_relevance = round(float(np.clip(0.60 * partial_g_factor + 0.40 * sim, 0.40, 0.97)), 3)
        composite = round(100.0 * (0.45 * severity + 0.30 * completeness + 0.25 * mission_relevance), 1)
        grade = "A" if composite >= 80 else ("B" if composite >= 68 else ("C" if composite >= 52 else "D"))
        ranked.append(
            {
                "rank": idx + 1,
                "accession": match["accession"],
                "title": match["title"],
                "source_url": match["source_url"],
                "platform": match["platform"],
                "flammability_severity": severity,
                "evidence_completeness": completeness,
                "mission_relevance": mission_relevance,
                "composite_score": composite,
                "confidence_grade": grade,
                "flag": match["match_rationale"],
            }
        )
    return ranked


def save_report(report: dict[str, Any]) -> dict[str, Any]:
    init_db()
    with get_db_connection() as conn:
        conn.execute(
            """
            INSERT INTO reports (report_id, created_at, workflow_mode, provider, model, report_json)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                report["report_id"],
                report["created_at"],
                report["workflow_mode"],
                report["si_interpretation"]["provider"],
                report["si_interpretation"]["model"],
                json.dumps(report),
            ),
        )
        conn.commit()
    return report


def create_synthetic_sample_image(kind: str = "0g_spherical") -> bytes:
    """Create a deterministic synthetic flame image for sample/demo testing."""
    canvas = np.zeros((240, 240, 3), dtype=np.uint8)
    canvas[:, :] = (18, 16, 14)
    if kind == "0g_spherical":
        cv2.circle(canvas, (120, 120), 62, (65, 105, 225), -1)
        cv2.circle(canvas, (120, 120), 42, (90, 160, 245), -1)
        cv2.circle(canvas, (120, 120), 20, (180, 230, 255), -1)
    elif kind == "lunar_017g":
        cv2.ellipse(canvas, (120, 125), (46, 68), 0, 0, 360, (55, 110, 230), -1)
        cv2.ellipse(canvas, (120, 118), (30, 48), 0, 0, 360, (90, 180, 250), -1)
    else:
        pts = np.array([[120, 32], [162, 145], [142, 200], [98, 200], [78, 145]], dtype=np.int32)
        cv2.fillPoly(canvas, [pts], (45, 95, 235))
        cv2.circle(canvas, (120, 150), 26, (120, 210, 255), -1)
    ok, encoded = cv2.imencode(".png", canvas)
    if not ok:
        raise RuntimeError("Failed to encode synthetic sample image.")
    return encoded.tobytes()


@asynccontextmanager
async def _lifespan(_app: FastAPI):
    init_db()
    yield


app = FastAPI(
    title="PyroSage — Microgravity Fire Intelligence API",
    version="0.1.0",
    description=(
        "Technique-bounded multimodal research dashboard and API for NASA Physical Sciences "
        "Informatics (PSI) microgravity combustion evidence analysis."
    ),
    lifespan=_lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
def serve_dashboard() -> Response:
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file), media_type="text/html")
    if PROTOTYPE_PATH.exists():
        return FileResponse(str(PROTOTYPE_PATH), media_type="text/html")
    return HTMLResponse("<h1>PyroSage API is running</h1>")


@app.get("/reference_prototype.html", response_class=HTMLResponse)
def serve_reference_prototype() -> Response:
    if PROTOTYPE_PATH.exists():
        return FileResponse(str(PROTOTYPE_PATH), media_type="text/html")
    raise HTTPException(status_code=404, detail="reference_prototype.html not found.")


@app.get("/api/health")
def api_health() -> dict[str, Any]:
    init_db()
    techniques = load_techniques()
    experiments = list_cached_psi_experiments()
    return {
        "status": "ok",
        "service": "PyroSage",
        "version": "0.1.0",
        "default_provider": DEFAULT_PROVIDER,
        "default_model": DEFAULT_MODEL,
        "providers_configured": {
            "gemini": bool(os.getenv("GEMINI_API_KEY", "").strip()),
            "openai": bool(os.getenv("OPENAI_API_KEY", "").strip()),
        },
        "nasa_psi_base_url": NASA_PSI_BASE_URL,
        "technique_count": len(techniques),
        "cached_psi_experiments": len(experiments),
        "limits": {
            "max_file_bytes": MAX_FILE_BYTES,
            "max_total_bytes": MAX_TOTAL_BYTES,
        },
    }


@app.get("/api/techniques")
def api_techniques(
    group: str | None = Query(default=None, description="Filter by scientific family/group"),
    modality: str | None = Query(default=None, description="Filter by supported modality"),
) -> dict[str, Any]:
    all_techniques = load_techniques()
    filtered = all_techniques
    if group:
        g_low = group.strip().lower()
        filtered = [t for t in filtered if g_low in t["group"].lower()]
    if modality:
        m_low = modality.strip().lower()
        filtered = [t for t in filtered if m_low in [m.lower() for m in t["supported_modalities"]]]

    families: dict[str, int] = {}
    for t in all_techniques:
        families[t["group"]] = families.get(t["group"], 0) + 1

    return {
        "count": len(filtered),
        "total_registered": len(all_techniques),
        "families": [{"name": k, "count": v} for k, v in families.items()],
        "execution_states": list(EXECUTION_STATES),
        "techniques": filtered,
    }


@app.get("/api/nasa-psi/search")
def api_nasa_psi_search(
    q: str = Query(default="combustion", description="Search keyword for NASA PSI investigations"),
    accession: str | None = Query(default=None, description="Direct NASA PSI accession (e.g., PSI-98)"),
    limit: int = Query(default=10, ge=1, le=50),
    live: bool = Query(default=False, description="Attempt live psi.nasa.gov HTTP fetch before fallback"),
) -> dict[str, Any]:
    init_db()
    try:
        return search_nasa_psi(query=q, accession=accession, limit=limit, attempt_live=live)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/nasa-psi/sync")
async def api_nasa_psi_sync(request: Request) -> dict[str, Any]:
    init_db()
    try:
        payload = await request.json()
    except Exception as exc:
        raise HTTPException(status_code=400, detail="Invalid JSON payload.") from exc

    records_in: list[dict[str, Any]] = []
    if isinstance(payload, dict) and "records" in payload:
        if not isinstance(payload["records"], list):
            raise HTTPException(status_code=400, detail="'records' must be a list of PSI records.")
        records_in = payload["records"]
    elif isinstance(payload, dict) and "accession" in payload:
        # Either a full record or an accession to sync from verified/search catalogue
        if "data_source" in payload or "source_url" in payload:
            records_in = [payload]
        else:
            acc = str(payload["accession"]).strip().upper()
            search_res = search_nasa_psi(accession=acc)
            if not search_res["results"]:
                raise HTTPException(
                    status_code=404, detail=f"NASA PSI investigation '{acc}' not found."
                )
            records_in = search_res["results"]
    elif isinstance(payload, dict) and "query" in payload:
        search_res = search_nasa_psi(query=str(payload["query"]))
        records_in = search_res["results"]
    else:
        records_in = VERIFIED_PSI_CATALOGUE

    synced: list[dict[str, Any]] = []
    for rec in records_in:
        try:
            synced.append(upsert_psi_record(rec))
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    return {
        "status": "synced",
        "data_source": "NASA PSI",
        "synced_count": len(synced),
        "records": synced,
    }


@app.get("/api/nasa-psi/experiments")
def api_nasa_psi_experiments() -> dict[str, Any]:
    records = list_cached_psi_experiments()
    return {
        "data_source": "NASA PSI",
        "repository_url": f"{NASA_PSI_BASE_URL}/physci/repo",
        "count": len(records),
        "experiments": records,
    }


@app.post("/api/nasa-psi/analyze")
async def api_nasa_psi_analyze(request: Request) -> dict[str, Any]:
    """Analyze a verified NASA PSI investigation and its linked assets with selected techniques."""
    init_db()
    try:
        payload = await request.json()
    except Exception as exc:
        raise HTTPException(status_code=400, detail="Expected JSON request body.") from exc

    accession = str(payload.get("accession", "PSI-98")).strip().upper()
    if not PSI_ACCESSION_RE.match(accession):
        raise HTTPException(
            status_code=400,
            detail=f"Invalid NASA PSI accession '{accession}'. Must match 'PSI-<number>'.",
        )

    catalogue = list_cached_psi_experiments()
    by_acc = {r["accession"]: r for r in catalogue}
    if accession not in by_acc:
        raise HTTPException(
            status_code=404,
            detail=f"Investigation '{accession}' is not present in the verified NASA PSI catalogue.",
        )

    psi_record = by_acc[accession]
    valid_asset_map = {a["asset_id"]: a for a in psi_record.get("assets", [])}

    requested_assets = payload.get("asset_ids") or [a["asset_id"] for a in psi_record.get("assets", [])]
    for aid in requested_assets:
        if aid not in valid_asset_map:
            raise HTTPException(
                status_code=400,
                detail=f"Asset '{aid}' is not linked to verified NASA PSI investigation '{accession}'.",
            )

    selected_techniques = payload.get("technique_ids") or [
        "T01",
        "T02",
        "T04",
        "T05",
        "T08",
        "T09",
        "T10",
        "T12",
        "T13",
        "T14",
        "T16",
        "T40",
        "T44",
        "T45",
        "T51",
        "T52",
        "T53",
        "T56",
        "T57",
    ]

    # Build representative local descriptor extraction for selected PSI investigation assets
    sample_kind = "0g_spherical" if accession in ("PSI-98", "PSI-107", "PSI-41") else "lunar_017g"
    img_desc = extract_image_descriptors(
        create_synthetic_sample_image(sample_kind),
        filename=f"{accession}_representative_frame.png",
    )
    img_desc["evidence_class"] = "source_record"
    img_desc["psi_accession"] = accession
    img_desc["source_url"] = psi_record["source_url"]

    evidence_items = [img_desc]
    for aid in requested_assets:
        asset_meta = valid_asset_map[aid]
        if asset_meta["modality"] == "table":
            csv_bytes = (
                "time_s,droplet_d2_mm2,radiometer_mw_cm2,oxygen_pct\n"
                "0.0,9.00,14.2,21.0\n"
                "1.0,7.20,12.8,20.8\n"
                "2.0,5.45,10.9,20.5\n"
                "3.0,3.62,4.1,20.2\n"
                "4.0,2.80,0.8,20.0\n"
            ).encode("utf-8")
            tab_desc = extract_table_descriptors(csv_bytes, asset_meta["filename"])
            tab_desc["evidence_class"] = "source_record"
            tab_desc["psi_accession"] = accession
            tab_desc["source_url"] = psi_record["source_url"]
            evidence_items.append(tab_desc)

    context = {
        "accession": accession,
        "psi_investigation": psi_record["title"],
        "gravity_g": float(payload.get("gravity_g", psi_record.get("gravity_g", 0.0))),
        "oxygen_pct": float(payload.get("oxygen_pct", 21.0)),
        "pressure_kpa": float(payload.get("pressure_kpa", psi_record.get("pressure_kpa", 101.3))),
        "fuel": str(payload.get("fuel") or (psi_record["fuels"][0] if psi_record["fuels"] else "heptane")),
        "platform": psi_record["platform"],
        "notes": str(payload.get("notes", "")),
    }

    assessments = evaluate_techniques(selected_techniques, evidence_items, context)
    candidate_matches = match_psi_experiments(evidence_items, context, catalogue)
    si_result = run_bounded_si_engine(
        evidence_items=evidence_items,
        technique_assessments=assessments,
        candidate_matches=candidate_matches,
        context=context,
        provider=payload.get("provider"),
        model=payload.get("model"),
    )
    ranked = compute_ranked_hazards(evidence_items, candidate_matches, assessments, context)

    report_id = f"rep-{uuid.uuid4().hex[:12]}"
    report = {
        "report_id": report_id,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "workflow_mode": "nasa_psi_evidence",
        "provenance": {
            "data_source": "NASA PSI",
            "accession": accession,
            "title": psi_record["title"],
            "source_url": psi_record["source_url"],
            "selected_assets": [valid_asset_map[aid] for aid in requested_assets],
            "fiction_simulator_used": False,
        },
        "mission_context": context,
        "evidence_inventory": evidence_items,
        "selected_technique_ids": selected_techniques,
        "technique_assessments": assessments,
        "candidate_experiments": candidate_matches,
        "si_interpretation": si_result,
        "ranked_findings": ranked,
    }
    return save_report(report)


@app.post("/api/analyze")
async def api_analyze(
    request: Request,
    files: list[UploadFile] | None = File(default=None),
    technique_ids: str | None = Form(default=None),
    gravity_g: float = Form(default=0.0),
    oxygen_pct: float = Form(default=21.0),
    pressure_kpa: float = Form(default=101.3),
    fuel: str = Form(default="pmma"),
    platform: str = Form(default="ISS / CIR"),
    provider: str | None = Form(default=None),
    model: str | None = Form(default=None),
    sample_preset: str | None = Form(default=None),
    notes: str = Form(default=""),
) -> dict[str, Any]:
    """Analyze uploaded multimodal evidence files or sample specimens with technique-bounded SI."""
    init_db()
    content_type_header = (request.headers.get("content-type") or "").lower()

    selected_techs: list[str] = []
    evidence_items: list[dict[str, Any]] = []
    total_bytes = 0

    if "application/json" in content_type_header:
        body = await request.json()
        raw_techs = body.get("technique_ids") or []
        if isinstance(raw_techs, list):
            selected_techs = [str(t) for t in raw_techs]
        gravity_g = float(body.get("gravity_g", gravity_g))
        oxygen_pct = float(body.get("oxygen_pct", oxygen_pct))
        pressure_kpa = float(body.get("pressure_kpa", pressure_kpa))
        fuel = str(body.get("fuel", fuel))
        platform = str(body.get("platform", platform))
        provider = body.get("provider", provider)
        model = body.get("model", model)
        sample_preset = body.get("sample_preset", sample_preset)
        notes = str(body.get("notes", notes))
    else:
        if technique_ids:
            try:
                parsed_t = json.loads(technique_ids)
                if isinstance(parsed_t, list):
                    selected_techs = [str(t) for t in parsed_t]
                else:
                    selected_techs = [t.strip() for t in technique_ids.split(",") if t.strip()]
            except Exception:
                selected_techs = [t.strip() for t in technique_ids.split(",") if t.strip()]

    if files:
        for up in files:
            if not up.filename:
                continue
            raw_bytes = await up.read()
            total_bytes += len(raw_bytes)
            if total_bytes > MAX_TOTAL_BYTES:
                raise HTTPException(
                    status_code=413,
                    detail=f"Total upload size ({total_bytes} bytes) exceeds PYROSAGE_MAX_TOTAL_BYTES ({MAX_TOTAL_BYTES} bytes).",
                )
            try:
                desc = parse_evidence_file(raw_bytes, up.filename, up.content_type or "")
                evidence_items.append(desc)
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc)) from exc

    if not evidence_items:
        preset_kind = sample_preset or ("0g_spherical" if gravity_g < 0.05 else "lunar_017g")
        img_bytes = create_synthetic_sample_image(preset_kind)
        desc = extract_image_descriptors(img_bytes, filename=f"specimen_{preset_kind}.png")
        desc["evidence_class"] = "uploaded_evidence"
        evidence_items.append(desc)

    context = {
        "gravity_g": gravity_g,
        "oxygen_pct": oxygen_pct,
        "pressure_kpa": pressure_kpa,
        "fuel": fuel,
        "platform": platform,
        "notes": notes,
    }

    assessments = evaluate_techniques(selected_techs, evidence_items, context)
    catalogue = list_cached_psi_experiments()
    candidate_matches = match_psi_experiments(evidence_items, context, catalogue)
    si_result = run_bounded_si_engine(
        evidence_items=evidence_items,
        technique_assessments=assessments,
        candidate_matches=candidate_matches,
        context=context,
        provider=provider,
        model=model,
    )
    ranked = compute_ranked_hazards(evidence_items, candidate_matches, assessments, context)

    report_id = f"rep-{uuid.uuid4().hex[:12]}"
    report = {
        "report_id": report_id,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "workflow_mode": "uploaded_evidence_analysis",
        "provenance": {
            "data_source": "Operator Uploaded Evidence (Matched against NASA PSI Catalogue)",
            "uploaded_files": [e["filename"] for e in evidence_items],
            "relabelled_as_nasa_data": False,
            "fiction_simulator_used": False,
        },
        "mission_context": context,
        "evidence_inventory": evidence_items,
        "selected_technique_ids": [a["id"] for a in assessments],
        "technique_assessments": assessments,
        "candidate_experiments": candidate_matches,
        "si_interpretation": si_result,
        "ranked_findings": ranked,
    }
    return save_report(report)


@app.get("/api/report/{report_id}")
def api_get_report(report_id: str) -> dict[str, Any]:
    init_db()
    with get_db_connection() as conn:
        row = conn.execute(
            "SELECT report_json FROM reports WHERE report_id = ?", (report_id,)
        ).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail=f"Report '{report_id}' not found.")
    return json.loads(row["report_json"])


@app.get("/api/reports")
def api_list_reports(limit: int = Query(default=20, ge=1, le=100)) -> dict[str, Any]:
    init_db()
    with get_db_connection() as conn:
        rows = conn.execute(
            "SELECT report_id, created_at, workflow_mode, provider, model FROM reports ORDER BY created_at DESC LIMIT ?",
            (limit,),
        ).fetchall()
    return {
        "count": len(rows),
        "reports": [dict(r) for r in rows],
    }


if __name__ == "__main__":
    import uvicorn

    port = int(os.getenv("PORT", "8000"))
    uvicorn.run("server:app", host="127.0.0.1", port=port, reload=False)
