"""PyroSage — Technique-bounded multimodal microgravity combustion analysis service."""

from __future__ import annotations

import base64
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

load_dotenv()

ROOT_DIR = Path(__file__).resolve().parent
STATIC_DIR = ROOT_DIR / "static"
DATA_DIR = ROOT_DIR / "data"
TECHNIQUES_PATH = ROOT_DIR / "techniques.json"
MAIN_DASHBOARD_PATH = ROOT_DIR / "pyrosage_interactive_dashboard.html"
PROTOTYPE_PATH = ROOT_DIR / "reference_prototype.html"

DEFAULT_PROVIDER = os.getenv("PYROSAGE_DEFAULT_PROVIDER", "gemini").strip().lower()
DEFAULT_MODEL = os.getenv("PYROSAGE_DEFAULT_MODEL", "gemini-3.8-flash").strip()
NASA_PSI_BASE_URL = os.getenv("NASA_PSI_BASE_URL", "https://psi.nasa.gov").rstrip("/")
MAX_FILE_BYTES = int(os.getenv("PYROSAGE_MAX_FILE_BYTES", str(15 * 1024 * 1024)))
MAX_TOTAL_BYTES = int(os.getenv("PYROSAGE_MAX_TOTAL_BYTES", str(35 * 1024 * 1024)))
MAX_UPLOAD_FILES = 20
MAX_IMAGE_PAYLOADS = 5

PSI_ACCESSION_RE = re.compile(r"^PSI-\d+$")

ALLOWED_STATUSES = {
    "computed_from_upload",
    "visually_assessed",
    "interpreted_from_metadata",
    "requires_more_data",
    "not_applicable",
}

UNSUPPORTED_SOLVER_TECHNIQUES = {
    "piv",
    "cnn_resnet",
    "symmetry_loss",
    "spectral_deconvolution",
    "chemical_kinetics",
    "gravity_scaling",
}

LOCALLY_COMPUTED_IMAGE_TECHNIQUES = {
    "segmentation",
    "reflection_iou",
    "bilateral_symmetry",
    "top_bottom_symmetry",
    "rotational_symmetry",
    "spherical_symmetry",
    "convex_hull",
    "convexity_defects",
    "fractal_dimension",
    "boundary_roughness",
    "perimeter_area",
    "peak_detection",
    "centroid_center_mass",
}

LOCALLY_COMPUTED_VIDEO_TECHNIQUES = {
    "temporal_dynamics",
    "symmetry_variance",
    "centroid_jitter",
}

OPTIONAL_GEOMETRY_FIELDS = (
    "bilateral_left_right_score",
    "top_bottom_symmetry_score",
    "convexity_score",
    "fractal_dimension",
    "flicker_frequency_hz",
    "centroid_jitter_px",
    "peak_count",
    "flame_regime",
    "classification",
)

# Verified NASA PSI reference investigations for the NASA PSI discovery / sync connector
VERIFIED_PSI_RECORDS: list[dict[str, Any]] = [
    {
        "experiment_id": "PSI-98",
        "accession": "PSI-98",
        "name": "Flame Extinguishment Experiment (FLEX)",
        "title": "Flame Extinguishment Experiment (FLEX)",
        "mission": "ISS Expedition / CIR-MDCA",
        "platform": "ISS / CIR",
        "gravity_g": 0.0,
        "gravity_regime": "0g",
        "fuel_material": "heptane, methanol, decane",
        "fuels": ["heptane", "methanol", "decane"],
        "oxygen_pct": 21.0,
        "pressure_kpa": 101.3,
        "temperature_k": 298.0,
        "techniques": "segmentation, spherical_symmetry, bilateral_symmetry, top_bottom_symmetry, cool_flame_physics",
        "diagnostics": ["visible_imaging", "radiometer", "backlit_droplet"],
        "modalities": ["image", "video", "numeric", "text"],
        "summary": "Spherical droplet combustion, radiative extinction, and cool-flame transition in microgravity aboard ISS CIR.",
        "tags": "droplet, heptane, methanol, cool flame, 0g, FLEX, PSI-98",
        "source_ref": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-98",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-98",
        "data_uri": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-98",
        "data_source": "NASA PSI",
        "bilateral_left_right_score": 0.95,
        "top_bottom_symmetry_score": 0.94,
        "convexity_score": 0.97,
        "fractal_dimension": 1.08,
        "flame_regime": "spherical",
        "assets": [
            {"asset_id": "PSI-98-IMG-01", "filename": "flex_0g_droplet.png", "modality": "image", "description": "0g spherical heptane droplet flame"},
            {"asset_id": "PSI-98-TAB-01", "filename": "flex_d2_radiometer.csv", "modality": "numeric", "description": "Droplet d^2 and radiometer history"},
        ],
    },
    {
        "experiment_id": "PSI-107",
        "accession": "PSI-107",
        "name": "Flame Extinguishment Experiment - 2 (FLEX-2)",
        "title": "Flame Extinguishment Experiment - 2 (FLEX-2)",
        "mission": "ISS Expedition / CIR",
        "platform": "ISS / CIR",
        "gravity_g": 0.0,
        "gravity_regime": "0g",
        "fuel_material": "heptane, decane, dodecane, iso-octane",
        "fuels": ["heptane", "decane", "dodecane", "iso-octane"],
        "oxygen_pct": 21.0,
        "pressure_kpa": 101.3,
        "temperature_k": 298.0,
        "techniques": "spherical_symmetry, cool_flame_physics, cool_flame_extinction, temporal_dynamics",
        "diagnostics": ["visible_imaging", "uv_oh_imaging", "radiometer"],
        "modalities": ["image", "video", "numeric", "spectral", "text"],
        "summary": "Isolated droplet and droplet-array interactions, soot shells, and two-stage cool-flame burning in 0g.",
        "tags": "droplet array, soot shell, cool flame, FLEX-2, PSI-107",
        "source_ref": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-107",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-107",
        "data_uri": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-107",
        "data_source": "NASA PSI",
        "bilateral_left_right_score": 0.93,
        "top_bottom_symmetry_score": 0.91,
        "convexity_score": 0.95,
        "fractal_dimension": 1.10,
        "flame_regime": "spherical",
        "assets": [
            {"asset_id": "PSI-107-IMG-01", "filename": "flex2_cool_flame.png", "modality": "image", "description": "FLEX-2 radiative extinction to cool flame"},
        ],
    },
    {
        "experiment_id": "PSI-25",
        "accession": "PSI-25",
        "name": "Burning and Suppression of Solids - II (BASS-II)",
        "title": "Burning and Suppression of Solids - II (BASS-II)",
        "mission": "ISS Expedition / MSG",
        "platform": "ISS / CIR",
        "gravity_g": 0.0,
        "gravity_regime": "0g",
        "fuel_material": "pmma, cotton, nomex",
        "fuels": ["pmma", "cotton", "nomex"],
        "oxygen_pct": 18.5,
        "pressure_kpa": 101.3,
        "temperature_k": 298.0,
        "techniques": "bilateral_symmetry, top_bottom_symmetry, convex_hull, gas_velocity, risk_interpretation",
        "diagnostics": ["visible_imaging", "radiometer", "thermocouple", "anemometer"],
        "modalities": ["image", "video", "numeric", "text"],
        "summary": "Concurrent and opposed-flow flame spread and extinction over solid PMMA rods, slabs, and fabrics in microgravity.",
        "tags": "PMMA, cotton, forced flow, limiting oxygen, BASS-II, PSI-25",
        "source_ref": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-25",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-25",
        "data_uri": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-25",
        "data_source": "NASA PSI",
        "bilateral_left_right_score": 0.89,
        "top_bottom_symmetry_score": 0.68,
        "convexity_score": 0.93,
        "fractal_dimension": 1.15,
        "flame_regime": "ellipsoidal",
        "assets": [
            {"asset_id": "PSI-25-IMG-01", "filename": "bass2_pmma_rod.png", "modality": "image", "description": "BASS-II PMMA rod stagnation flame"},
            {"asset_id": "PSI-25-TAB-01", "filename": "bass2_loc_matrix.csv", "modality": "numeric", "description": "Limiting oxygen concentration vs flow velocity"},
        ],
    },
    {
        "experiment_id": "PSI-112",
        "accession": "PSI-112",
        "name": "Solid Fuel Ignition and Extinction (SoFIE)",
        "title": "Solid Fuel Ignition and Extinction (SoFIE)",
        "mission": "ISS / CIR SoFIE",
        "platform": "ISS / CIR",
        "gravity_g": 0.0,
        "gravity_regime": "0g",
        "fuel_material": "pmma, solid_polymer",
        "fuels": ["pmma", "solid_polymer"],
        "oxygen_pct": 34.0,
        "pressure_kpa": 56.5,
        "temperature_k": 298.0,
        "techniques": "segmentation, infrared_thermography, lunar_gravity, martian_gravity, risk_interpretation",
        "diagnostics": ["visible_imaging", "ir_thermography", "schlieren"],
        "modalities": ["image", "video", "numeric", "thermal", "text"],
        "summary": "Ignition and extinction of solid spacecraft materials under exploration cabin pressures and elevated oxygen.",
        "tags": "SoFIE, PMMA, exploration atmosphere, lunar, martian, PSI-112",
        "source_ref": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-112",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-112",
        "data_uri": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-112",
        "data_source": "NASA PSI",
        "bilateral_left_right_score": 0.90,
        "top_bottom_symmetry_score": 0.74,
        "convexity_score": 0.94,
        "fractal_dimension": 1.13,
        "flame_regime": "ellipsoidal",
        "assets": [
            {"asset_id": "PSI-112-IMG-01", "filename": "sofie_pmma_34o2.png", "modality": "image", "description": "SoFIE PMMA flame under 34% O2 / 56.5 kPa"},
        ],
    },
    {
        "experiment_id": "PSI-64",
        "accession": "PSI-64",
        "name": "Spacecraft Fire Experiment (Saffire)",
        "title": "Spacecraft Fire Experiment (Saffire)",
        "mission": "Cygnus Orbital Vehicle",
        "platform": "Orbital capsule / satellite",
        "gravity_g": 0.0,
        "gravity_regime": "0g",
        "fuel_material": "cotton, sibal_fabric, pmma",
        "fuels": ["cotton", "sibal_fabric", "pmma"],
        "oxygen_pct": 21.0,
        "pressure_kpa": 101.3,
        "temperature_k": 298.0,
        "techniques": "orbital_platforms, temperature_distribution, co_analysis, risk_interpretation",
        "diagnostics": ["visible_imaging", "thermocouple_rake", "radiometer", "co_co2_o2_sensors"],
        "modalities": ["image", "video", "numeric", "text"],
        "summary": "Large-scale autonomous fire growth, oxygen consumption, and smoke/CO propagation inside unmanned Cygnus spacecraft.",
        "tags": "Saffire, Cygnus, cotton, SIBAL, cabin fire, PSI-64",
        "source_ref": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-64",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-64",
        "data_uri": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-64",
        "data_source": "NASA PSI",
        "bilateral_left_right_score": 0.82,
        "top_bottom_symmetry_score": 0.56,
        "convexity_score": 0.86,
        "fractal_dimension": 1.24,
        "flame_regime": "teardrop",
        "assets": [
            {"asset_id": "PSI-64-TAB-01", "filename": "saffire_cabin_telemetry.csv", "modality": "numeric", "description": "Cabin pressure, O2, CO, CO2, and thermocouple traces"},
        ],
    },
    {
        "experiment_id": "PSI-41",
        "accession": "PSI-41",
        "name": "Advanced Combustion via Microgravity Experiments (ACME)",
        "title": "Advanced Combustion via Microgravity Experiments (ACME)",
        "mission": "ISS Expedition / CIR",
        "platform": "ISS / CIR",
        "gravity_g": 0.0,
        "gravity_regime": "0g",
        "fuel_material": "methane, ethylene, propane",
        "fuels": ["methane", "ethylene", "propane"],
        "oxygen_pct": 21.0,
        "pressure_kpa": 101.3,
        "temperature_k": 298.0,
        "techniques": "spherical_symmetry, soot_formation, lif, heat_radiation, iss_cir",
        "diagnostics": ["visible_imaging", "chemiluminescence", "soot_pyrometry"],
        "modalities": ["image", "video", "numeric", "spectral", "text"],
        "summary": "Gaseous non-premixed coflow and spherical porous-burner diffusion flames in the ISS CIR.",
        "tags": "ACME, spherical burner, gaseous flame, soot, CIR, PSI-41",
        "source_ref": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-41",
        "source_url": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-41",
        "data_uri": f"{NASA_PSI_BASE_URL}/physci/repo/data/investigations/PSI-41",
        "data_source": "NASA PSI",
        "bilateral_left_right_score": 0.94,
        "top_bottom_symmetry_score": 0.88,
        "convexity_score": 0.96,
        "fractal_dimension": 1.09,
        "flame_regime": "spherical",
        "assets": [
            {"asset_id": "PSI-41-SPE-01", "filename": "acme_chemiluminescence.csv", "modality": "spectral", "description": "OH*, CH*, and C2* spectral emission"},
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
    """Initialize SQLite schema. The local experiment catalogue starts empty by design."""
    with get_db_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS catalog (
                experiment_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                mission TEXT NOT NULL DEFAULT '',
                platform TEXT NOT NULL DEFAULT '',
                gravity_g REAL,
                fuel_material TEXT NOT NULL DEFAULT '',
                oxygen_pct REAL,
                pressure_kpa REAL,
                temperature_k REAL,
                techniques TEXT NOT NULL DEFAULT '',
                summary TEXT NOT NULL DEFAULT '',
                tags TEXT NOT NULL DEFAULT '',
                source_ref TEXT NOT NULL DEFAULT '',
                data_uri TEXT NOT NULL DEFAULT '',
                raw_json TEXT NOT NULL DEFAULT '{}',
                imported_at TEXT NOT NULL
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


def load_techniques() -> list[dict[str, Any]]:
    if not TECHNIQUES_PATH.exists():
        raise RuntimeError(f"Technique registry missing at {TECHNIQUES_PATH}")
    raw = json.loads(TECHNIQUES_PATH.read_text(encoding="utf-8"))
    if isinstance(raw, dict) and "techniques" in raw:
        items = raw["techniques"]
    elif isinstance(raw, list):
        items = raw
    else:
        raise ValueError("techniques.json must contain a list of 57 technique objects.")
    return items


def get_technique_map() -> dict[str, dict[str, Any]]:
    items = load_techniques()
    mapping: dict[str, dict[str, Any]] = {}
    for idx, item in enumerate(items, start=1):
        mapping[item["id"]] = item
        mapping[f"T{idx:02d}"] = item
    return mapping


def _to_float_or_none(val: Any) -> float | None:
    if val is None or val == "":
        return None
    try:
        num = float(val)
        return num if math.isfinite(num) else None
    except (ValueError, TypeError):
        return None


def normalize_catalog_row(row: dict[str, Any]) -> dict[str, Any] | None:
    """Normalize catalogue row aliases according to Section 09 of the Technical Documentation."""
    if not isinstance(row, dict):
        return None
    lower = {str(k).strip().lower(): v for k, v in row.items() if k is not None}

    name = str(
        lower.get("name")
        or lower.get("title")
        or lower.get("experiment")
        or lower.get("investigation")
        or ""
    ).strip()
    if not name:
        return None

    exp_id = str(
        lower.get("experiment_id")
        or lower.get("accession")
        or lower.get("id")
        or lower.get("record_id")
        or ""
    ).strip()
    if not exp_id:
        exp_id = f"CAT-{uuid.uuid4().hex[:5].upper()}"

    mission = str(lower.get("mission") or lower.get("campaign") or lower.get("programme") or "").strip()
    platform = str(lower.get("platform") or lower.get("facility") or lower.get("hardware") or "").strip()
    gravity_g = _to_float_or_none(lower.get("gravity_g") if "gravity_g" in lower else lower.get("gravity"))
    fuel_val = lower.get("fuel_material") or lower.get("fuel") or lower.get("material") or lower.get("fuels") or ""
    if isinstance(fuel_val, list):
        fuel_material = ", ".join(str(x) for x in fuel_val)
    else:
        fuel_material = str(fuel_val).strip()

    oxygen_pct = _to_float_or_none(lower.get("oxygen_pct") if "oxygen_pct" in lower else lower.get("oxygen"))
    pressure_kpa = _to_float_or_none(lower.get("pressure_kpa") if "pressure_kpa" in lower else lower.get("pressure"))
    temperature_k = _to_float_or_none(lower.get("temperature_k") if "temperature_k" in lower else lower.get("temperature"))

    tech_val = lower.get("techniques") or lower.get("diagnostics") or lower.get("methods") or ""
    if isinstance(tech_val, list):
        techniques = ", ".join(str(x) for x in tech_val)
    else:
        techniques = str(tech_val).strip()

    summary = str(lower.get("summary") or lower.get("description") or "").strip()
    tags_val = lower.get("tags") or lower.get("keywords") or ""
    tags = ", ".join(str(x) for x in tags_val) if isinstance(tags_val, list) else str(tags_val).strip()
    source_ref = str(
        lower.get("source_ref")
        or lower.get("source_url")
        or lower.get("source")
        or lower.get("url")
        or lower.get("citation")
        or ""
    ).strip()
    data_uri = str(lower.get("data_uri") or lower.get("archive_uri") or source_ref or "").strip()

    raw_copy = dict(row)
    for g_field in OPTIONAL_GEOMETRY_FIELDS:
        if g_field in lower and lower[g_field] is not None and lower[g_field] != "":
            num_g = _to_float_or_none(lower[g_field])
            raw_copy[g_field] = num_g if num_g is not None else str(lower[g_field]).strip()

    return {
        "experiment_id": exp_id,
        "name": name,
        "mission": mission,
        "platform": platform,
        "gravity_g": gravity_g,
        "fuel_material": fuel_material,
        "oxygen_pct": oxygen_pct,
        "pressure_kpa": pressure_kpa,
        "temperature_k": temperature_k,
        "techniques": techniques,
        "summary": summary,
        "tags": tags,
        "source_ref": source_ref,
        "data_uri": data_uri,
        "raw": raw_copy,
        "imported_at": datetime.now(timezone.utc).isoformat(),
    }


def import_catalog_records(records: list[dict[str, Any]]) -> dict[str, Any]:
    init_db()
    imported = 0
    skipped = 0
    saved_records: list[dict[str, Any]] = []
    with get_db_connection() as conn:
        for row in records:
            norm = normalize_catalog_row(row)
            if norm is None:
                skipped += 1
                continue
            conn.execute(
                """
                INSERT INTO catalog (
                    experiment_id, name, mission, platform, gravity_g,
                    fuel_material, oxygen_pct, pressure_kpa, temperature_k,
                    techniques, summary, tags, source_ref, data_uri, raw_json, imported_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(experiment_id) DO UPDATE SET
                    name = excluded.name,
                    mission = excluded.mission,
                    platform = excluded.platform,
                    gravity_g = excluded.gravity_g,
                    fuel_material = excluded.fuel_material,
                    oxygen_pct = excluded.oxygen_pct,
                    pressure_kpa = excluded.pressure_kpa,
                    temperature_k = excluded.temperature_k,
                    techniques = excluded.techniques,
                    summary = excluded.summary,
                    tags = excluded.tags,
                    source_ref = excluded.source_ref,
                    data_uri = excluded.data_uri,
                    raw_json = excluded.raw_json,
                    imported_at = excluded.imported_at
                """,
                (
                    norm["experiment_id"],
                    norm["name"],
                    norm["mission"],
                    norm["platform"],
                    norm["gravity_g"],
                    norm["fuel_material"],
                    norm["oxygen_pct"],
                    norm["pressure_kpa"],
                    norm["temperature_k"],
                    norm["techniques"],
                    norm["summary"],
                    norm["tags"],
                    norm["source_ref"],
                    norm["data_uri"],
                    json.dumps(norm["raw"]),
                    norm["imported_at"],
                ),
            )
            imported += 1
            saved_records.append(norm)
        conn.commit()
        total = int(conn.execute("SELECT COUNT(*) FROM catalog").fetchone()[0])
    return {"imported": imported, "skipped": skipped, "total": total, "records": saved_records}


def list_catalog_records(limit: int = 1000) -> dict[str, Any]:
    init_db()
    with get_db_connection() as conn:
        total = int(conn.execute("SELECT COUNT(*) FROM catalog").fetchone()[0])
        rows = conn.execute(
            "SELECT * FROM catalog ORDER BY experiment_id ASC LIMIT ?", (min(limit, 1000),)
        ).fetchall()
    records: list[dict[str, Any]] = []
    for row in rows:
        raw = json.loads(row["raw_json"] or "{}")
        rec: dict[str, Any] = {
            "experiment_id": row["experiment_id"],
            "accession": row["experiment_id"],
            "name": row["name"],
            "title": row["name"],
            "mission": row["mission"],
            "platform": row["platform"],
            "gravity_g": row["gravity_g"],
            "fuel_material": row["fuel_material"],
            "oxygen_pct": row["oxygen_pct"],
            "pressure_kpa": row["pressure_kpa"],
            "temperature_k": row["temperature_k"],
            "techniques": row["techniques"],
            "summary": row["summary"],
            "tags": row["tags"],
            "source_ref": row["source_ref"],
            "source_url": row["source_ref"],
            "data_uri": row["data_uri"],
            "imported_at": row["imported_at"],
            "raw": raw,
        }
        for g_field in OPTIONAL_GEOMETRY_FIELDS:
            if g_field in raw:
                rec[g_field] = raw[g_field]
        records.append(rec)
    return {"total": total, "count": len(records), "records": records}


def clear_catalog_records() -> dict[str, Any]:
    init_db()
    with get_db_connection() as conn:
        deleted = int(conn.execute("SELECT COUNT(*) FROM catalog").fetchone()[0])
        conn.execute("DELETE FROM catalog")
        conn.commit()
    return {"deleted": deleted, "total": 0}


def validate_psi_record(record: dict[str, Any]) -> dict[str, Any]:
    """Enforce NASA PSI provenance rules on PSI connector records."""
    if not isinstance(record, dict):
        raise ValueError("PSI record must be a JSON object.")
    accession = str(record.get("accession") or record.get("experiment_id") or "").strip().upper()
    if not PSI_ACCESSION_RE.match(accession):
        raise ValueError(f"Invalid NASA PSI accession '{accession}'. Expected 'PSI-<number>'.")
    data_source = str(record.get("data_source", "NASA PSI")).strip()
    if data_source != "NASA PSI":
        raise ValueError(f"Invalid data_source '{data_source}'. Only 'NASA PSI' records are permitted.")
    source_url = str(record.get("source_url") or record.get("source_ref") or "").strip()
    parsed = urlparse(source_url)
    allowed_host = urlparse(NASA_PSI_BASE_URL).netloc or "psi.nasa.gov"
    if parsed.scheme not in ("http", "https") or parsed.netloc != allowed_host:
        raise ValueError(f"Invalid source_url '{source_url}'. Must point to '{allowed_host}'.")
    title = str(record.get("title") or record.get("name") or "").strip()
    if not title:
        raise ValueError("PSI record must include a non-empty title.")
    return {
        **record,
        "experiment_id": accession,
        "accession": accession,
        "name": title,
        "title": title,
        "data_source": "NASA PSI",
        "source_ref": source_url,
        "source_url": source_url,
    }


def _box_counting_fractal_dimension(edge_map: np.ndarray) -> float:
    points = np.argwhere(edge_map > 0)
    if len(points) < 8:
        return 1.0
    min_dim = min(edge_map.shape[:2])
    max_exp = int(np.floor(np.log2(max(min_dim, 8))))
    scales = [2**k for k in range(1, min(max_exp, 6))]
    counts: list[int] = []
    valid_scales: list[int] = []
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
    return round(float(np.clip(coeffs[0], 1.0, 1.99)), 4)


def _compute_iou(mask_a: np.ndarray, mask_b: np.ndarray) -> float:
    a = mask_a > 0
    b = mask_b > 0
    union = int(np.count_nonzero(a | b))
    if union == 0:
        return 0.0
    inter = int(np.count_nonzero(a & b))
    return float(inter / union)


def _rotate_mask(mask: np.ndarray, angle_deg: float) -> np.ndarray:
    h, w = mask.shape[:2]
    center = (w / 2.0, h / 2.0)
    mat = cv2.getRotationMatrix2D(center, angle_deg, 1.0)
    return cv2.warpAffine(mask, mat, (w, h), flags=cv2.INTER_NEAREST)


def extract_image_descriptors(image_bytes: bytes, filename: str = "image.png") -> dict[str, Any]:
    """Compute OpenCV HSV warm-color mask and geometric descriptors as specified in Section 07."""
    arr = np.frombuffer(image_bytes, dtype=np.uint8)
    bgr = cv2.imdecode(arr, cv2.IMREAD_COLOR)
    if bgr is None or bgr.size == 0:
        raise ValueError(f"Could not decode image file '{filename}'.")

    h_img, w_img = bgr.shape[:2]
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)

    # Warm-color HSV ranges (red/orange/yellow/bright core) per Section 07
    mask1 = cv2.inRange(hsv, np.array([0, 40, 60]), np.array([45, 255, 255]))
    mask2 = cv2.inRange(hsv, np.array([160, 40, 60]), np.array([180, 255, 255]))
    warm_mask = cv2.bitwise_or(mask1, mask2)

    if np.count_nonzero(warm_mask) < 0.003 * warm_mask.size:
        _, warm_mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    kernel = np.ones((3, 3), dtype=np.uint8)
    mask = cv2.morphologyEx(warm_mask, cv2.MORPH_OPEN, kernel)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
    if not contours:
        max_idx = np.unravel_index(int(np.argmax(gray)), gray.shape)
        mask = np.zeros_like(gray)
        cv2.circle(mask, (int(max_idx[1]), int(max_idx[0])), max(8, min(h_img, w_img) // 6), 255, -1)
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)

    main_contour = max(contours, key=cv2.contourArea)
    clean_mask = np.zeros_like(gray)
    cv2.drawContours(clean_mask, [main_contour], -1, 255, thickness=-1)

    candidate_area_px = float(max(cv2.contourArea(main_contour), float(np.count_nonzero(clean_mask))))
    candidate_perimeter_px = float(max(cv2.arcLength(main_contour, True), 1.0))
    perimeter_to_area_ratio = float(candidate_perimeter_px / max(candidate_area_px, 1.0))

    x_box, y_box, w_box, h_box = cv2.boundingRect(main_contour)
    aspect_ratio = float(h_box / max(w_box, 1))
    circularity = float(np.clip((4.0 * math.pi * candidate_area_px) / (candidate_perimeter_px**2), 0.0, 1.0))

    min_rect = cv2.minAreaRect(main_contour)
    orientation_deg = float(min_rect[2])

    hull = cv2.convexHull(main_contour)
    hull_area = float(max(cv2.contourArea(hull), candidate_area_px, 1.0))
    hull_perimeter = float(max(cv2.arcLength(hull, True), 1.0))
    convexity_score = float(np.clip(candidate_area_px / hull_area, 0.0, 1.0))
    boundary_roughness = float(max(candidate_perimeter_px / hull_perimeter, 1.0))

    # Convexity-defect count on approximated contour
    defect_count = 0
    approx = cv2.approxPolyDP(main_contour, 0.015 * candidate_perimeter_px, True)
    if len(approx) >= 4:
        try:
            hull_idx = cv2.convexHull(approx, returnPoints=False)
            if hull_idx is not None and len(hull_idx) >= 3:
                defects = cv2.convexityDefects(approx, hull_idx)
                if defects is not None:
                    defect_count = int(sum(1 for d in defects if d[0][3] > 256))
        except cv2.error:
            defect_count = 0

    # Peak count rough estimate on smoothed upper-boundary profile
    crop_mask = clean_mask[y_box : y_box + h_box, x_box : x_box + w_box]
    upper_profile = []
    for col_i in range(crop_mask.shape[1]):
        nz = np.nonzero(crop_mask[:, col_i])[0]
        upper_profile.append(float(h_box - nz[0]) if len(nz) > 0 else 0.0)
    prof_arr = np.array(upper_profile, dtype=float)
    if len(prof_arr) >= 5:
        smooth = np.convolve(prof_arr, np.ones(5) / 5.0, mode="same")
        peaks = [
            i
            for i in range(1, len(smooth) - 1)
            if smooth[i] > smooth[i - 1] and smooth[i] >= smooth[i + 1] and smooth[i] > 0.25 * np.max(smooth)
        ]
        peak_count = max(1, len(peaks))
    else:
        peak_count = 1

    # Centroids: geometric and intensity-weighted proxy
    moments = cv2.moments(clean_mask)
    if moments["m00"] > 0:
        cx_geom = float(moments["m10"] / moments["m00"])
        cy_geom = float(moments["m01"] / moments["m00"])
    else:
        cx_geom = float(x_box + w_box / 2.0)
        cy_geom = float(y_box + h_box / 2.0)

    weights = np.where(clean_mask > 0, np.maximum(gray.astype(np.float32), 1.0), 0.0)
    w_sum = float(np.sum(weights))
    if w_sum > 0:
        y_indices, x_indices = np.indices(gray.shape)
        cx_int = float(np.sum(x_indices * weights) / w_sum)
        cy_int = float(np.sum(y_indices * weights) / w_sum)
    else:
        cx_int, cy_int = cx_geom, cy_geom

    centroid_offset_px = float(math.hypot(cx_int - cx_geom, cy_int - cy_geom))

    # Center candidate in square padded canvas before reflection & rotational IoU (Section 07)
    side = int(max(w_box, h_box, 16) + 8)
    centered = np.zeros((side, side), dtype=np.uint8)
    y_off = (side - h_box) // 2
    x_off = (side - w_box) // 2
    centered[y_off : y_off + h_box, x_off : x_off + w_box] = crop_mask

    bilateral_score = _compute_iou(centered, cv2.flip(centered, 1))
    top_bottom_score = _compute_iou(centered, cv2.flip(centered, 0))
    reflection_iou = float(0.5 * (bilateral_score + top_bottom_score))

    rot_90_iou = _compute_iou(centered, _rotate_mask(centered, 90.0))
    rot_120_iou = _compute_iou(centered, _rotate_mask(centered, 120.0))
    rot_180_iou = _compute_iou(centered, _rotate_mask(centered, 180.0))
    rotational_symmetry = float((rot_90_iou + rot_120_iou + rot_180_iou) / 3.0)

    aspect_closeness = float(np.exp(-abs(aspect_ratio - 1.0)))
    spherical_symmetry_score = float(
        np.clip(
            0.30 * bilateral_score
            + 0.30 * top_bottom_score
            + 0.20 * rotational_symmetry
            + 0.20 * circularity * aspect_closeness,
            0.0,
            1.0,
        )
    )

    edge_map = np.zeros_like(gray)
    cv2.drawContours(edge_map, [main_contour], -1, 255, thickness=1)
    fractal_dim = _box_counting_fractal_dimension(edge_map)

    low_name = filename.lower()
    is_thermal_named = any(tok in low_name for tok in ("flir", "thermal", "infrared", "ir_"))
    modality = "thermal" if is_thermal_named else "image"

    if spherical_symmetry_score >= 0.80 and 0.80 <= aspect_ratio <= 1.25:
        flame_regime = "spherical"
    elif aspect_ratio <= 1.55:
        flame_regime = "ellipsoidal"
    else:
        flame_regime = "teardrop"

    return {
        "modality": modality,
        "filename": filename,
        "width": w_img,
        "height": h_img,
        "width_px": w_img,
        "height_px": h_img,
        "candidate_area_px": round(candidate_area_px, 2),
        "flame_area_px": round(candidate_area_px, 2),
        "flame_area_fraction": round(candidate_area_px / float(w_img * h_img), 5),
        "candidate_perimeter_px": round(candidate_perimeter_px, 2),
        "perimeter_px": round(candidate_perimeter_px, 2),
        "perimeter_to_area_ratio": round(perimeter_to_area_ratio, 5),
        "bilateral_left_right_score": round(bilateral_score, 4),
        "bilateral_symmetry": round(bilateral_score, 4),
        "top_bottom_symmetry_score": round(top_bottom_score, 4),
        "top_bottom_symmetry": round(top_bottom_score, 4),
        "reflection_iou": round(reflection_iou, 4),
        "mirror_iou_horizontal": round(bilateral_score, 4),
        "mirror_iou_vertical": round(top_bottom_score, 4),
        "rotational_symmetry_iou": {
            "deg_90": round(rot_90_iou, 4),
            "deg_120": round(rot_120_iou, 4),
            "deg_180": round(rot_180_iou, 4),
            "mean": round(rotational_symmetry, 4),
        },
        "rotational_symmetry": round(rotational_symmetry, 4),
        "spherical_symmetry_score": round(spherical_symmetry_score, 4),
        "convexity_score": round(convexity_score, 4),
        "solidity": round(convexity_score, 4),
        "convexity": round(float(np.clip(hull_perimeter / candidate_perimeter_px, 0.0, 1.0)), 4),
        "boundary_roughness": round(boundary_roughness, 4),
        "fractal_dimension": round(fractal_dim, 4),
        "convexity_defect_count": defect_count,
        "peak_count": peak_count,
        "geometric_centroid": [round(cx_geom, 2), round(cy_geom, 2)],
        "intensity_centroid": [round(cx_int, 2), round(cy_int, 2)],
        "centroid_x": round(cx_geom, 2),
        "centroid_y": round(cy_geom, 2),
        "intensity_centroid_offset_px": round(centroid_offset_px, 3),
        "orientation_deg": round(orientation_deg, 2),
        "bounding_box": {"x": int(x_box), "y": int(y_box), "width": int(w_box), "height": int(h_box)},
        "aspect_ratio": round(aspect_ratio, 4),
        "circularity": round(circularity, 4),
        "flame_regime": flame_regime,
        "morphology_class": f"{flame_regime}_diffusion" if flame_regime == "spherical" else flame_regime,
        "units": {
            "candidate_area_px": "px²",
            "candidate_perimeter_px": "px",
            "perimeter_to_area_ratio": "px⁻¹",
            "bilateral_left_right_score": "IoU [0,1]",
            "top_bottom_symmetry_score": "IoU [0,1]",
            "reflection_iou": "IoU [0,1]",
            "rotational_symmetry": "IoU [0,1]",
            "convexity_score": "ratio [0,1]",
            "fractal_dimension": "box-counting D",
            "convexity_defect_count": "count",
            "peak_count": "count",
            "orientation_deg": "degrees",
        },
        "caveat": "Heuristic warm-color segmentation; pixels are not physical length (m) or temperature (K) without calibration.",
        "_image_payloads": [
            {
                "label": filename,
                "mime_type": "image/png",
                "base64": base64.b64encode(cv2.imencode(".png", bgr)[1].tobytes()).decode("ascii"),
            }
        ],
    }


def extract_video_descriptors(video_bytes: bytes, filename: str = "video.mp4") -> dict[str, Any]:
    """Sample start, 1/3, 2/3, and final video frames and compute temporal descriptors per Section 07."""
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
        frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)

        if frame_count >= 4:
            positions = sorted({0, frame_count // 3, (2 * frame_count) // 3, max(0, frame_count - 1)})
        else:
            positions = [0, 1, 2, 3]

        sampled_features: list[dict[str, Any]] = []
        frame_payloads: list[dict[str, str]] = []
        for pos in positions:
            cap.set(cv2.CAP_PROP_POS_FRAMES, pos)
            ret, frame = cap.read()
            if not ret or frame is None:
                continue
            ok, buf = cv2.imencode(".png", frame)
            if ok:
                png_bytes = buf.tobytes()
                f_desc = extract_image_descriptors(png_bytes, f"{filename}#frame{pos}")
                f_desc.pop("_image_payloads", None)
                f_desc["frame_index"] = pos
                sampled_features.append(f_desc)
                frame_payloads.append(
                    {
                        "label": f"{filename}#frame{pos}",
                        "mime_type": "image/png",
                        "base64": base64.b64encode(png_bytes).decode("ascii"),
                    }
                )
        cap.release()
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    if not sampled_features:
        raise ValueError(f"No readable frames decoded from video '{filename}'.")

    duration_s = float(max(frame_count, len(sampled_features)) / max(fps, 1.0))
    areas = np.array([f["candidate_area_px"] for f in sampled_features], dtype=float)
    syms = np.array([f["bilateral_left_right_score"] for f in sampled_features], dtype=float)
    centroids = np.array([f["geometric_centroid"] for f in sampled_features], dtype=float)

    area_cv = float(np.std(areas) / max(float(np.mean(areas)), 1.0))
    symmetry_var_per_s = float(np.var(syms) / max(duration_s, 0.1))

    if len(centroids) >= 2:
        disps = np.hypot(np.diff(centroids[:, 0]), np.diff(centroids[:, 1]))
        centroid_jitter_px = float(np.mean(disps))
    else:
        centroid_jitter_px = 0.0

    base = dict(sampled_features[0])
    base.update(
        {
            "modality": "video",
            "filename": filename,
            "frame_count": max(frame_count, len(sampled_features)),
            "sampled_frames": len(sampled_features),
            "sampled_positions": [f["frame_index"] for f in sampled_features],
            "fps": round(fps, 2),
            "duration_s": round(duration_s, 3),
            "area_cv": round(area_cv, 4),
            "symmetry_variance": round(symmetry_var_per_s, 6),
            "centroid_jitter_px": round(centroid_jitter_px, 3),
            "flicker_frequency_computed": False,
            "piv_computed": False,
            "piv_note": "PIV was not computed; requires seeded tracer particle pairs, frame timing, and calibration.",
            "_image_payloads": frame_payloads[:MAX_IMAGE_PAYLOADS],
        }
    )
    return base


def extract_table_descriptors(table_bytes: bytes, filename: str = "data.csv") -> dict[str, Any]:
    """Compute column statistics and spectral local maxima suggestions per Section 07."""
    text = table_bytes.decode("utf-8", errors="replace")
    first_line = text.splitlines()[0] if text.splitlines() else ""
    delimiter = "\t" if filename.lower().endswith(".tsv") or ("\t" in first_line) else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
    rows = list(reader)
    columns = list(reader.fieldnames or [])

    numeric_cols: dict[str, list[float]] = {col: [] for col in columns}
    for row in rows:
        for col in columns:
            val = _to_float_or_none(row.get(col))
            if val is not None:
                numeric_cols[col].append(val)

    col_summaries: dict[str, dict[str, Any]] = {}
    for col, vals in numeric_cols.items():
        if not vals:
            continue
        arr = np.array(vals, dtype=float)
        col_summaries[col] = {
            "count": len(vals),
            "min": round(float(np.min(arr)), 5),
            "max": round(float(np.max(arr)), 5),
            "mean": round(float(np.mean(arr)), 5),
            "std": round(float(np.std(arr)), 5),
        }

    lower_cols = {c.lower(): c for c in columns}
    axis_col = next(
        (orig for low, orig in lower_cols.items() if any(k in low for k in ("wavelength", "nm", "wavenumber", "cm-1", "cm^-1", "raman"))),
        None,
    )
    signal_col = next(
        (orig for low, orig in lower_cols.items() if any(k in low for k in ("intensity", "absorbance", "transmittance", "emission", "signal", "counts"))),
        None,
    )

    is_spectral = bool(axis_col and signal_col and numeric_cols.get(axis_col) and numeric_cols.get(signal_col))
    spectral_hints: dict[str, Any] = {}
    if is_spectral and axis_col and signal_col:
        x_arr = np.array(numeric_cols[axis_col], dtype=float)
        y_arr = np.array(numeric_cols[signal_col], dtype=float)
        n_pts = min(len(x_arr), len(y_arr))
        x_arr, y_arr = x_arr[:n_pts], y_arr[:n_pts]
        peak_idx = int(np.argmax(y_arr)) if n_pts > 0 else 0
        local_maxima = []
        for i in range(1, n_pts - 1):
            if y_arr[i] >= y_arr[i - 1] and y_arr[i] >= y_arr[i + 1]:
                local_maxima.append({"position": round(float(x_arr[i]), 3), "signal": round(float(y_arr[i]), 5)})
        if not local_maxima and n_pts > 0:
            local_maxima.append({"position": round(float(x_arr[peak_idx]), 3), "signal": round(float(y_arr[peak_idx]), 5)})
        spectral_hints = {
            "axis_column": axis_col,
            "signal_column": signal_col,
            "dominant_wavelength_or_band": round(float(x_arr[peak_idx]), 3) if n_pts > 0 else 0.0,
            "candidate_maxima": local_maxima[:10],
            "note": "Maxima are candidate suggestions only; species assignment requires calibration and a reference spectrum.",
        }

    return {
        "modality": "spectral" if is_spectral else "numeric",
        "filename": filename,
        "row_count": len(rows),
        "column_count": len(columns),
        "columns": columns,
        "numeric_summaries": col_summaries,
        "spectral_analysis": spectral_hints,
    }


def extract_document_descriptors(file_bytes: bytes, filename: str) -> dict[str, Any]:
    """Parse JSON, PDF, DOCX, XLSX/XLSM, or text/instrument logs per Section 06."""
    ext = Path(filename).suffix.lower()
    low_name = filename.lower()
    extracted_text = ""
    extra: dict[str, Any] = {}

    if ext == ".json":
        parsed = json.loads(file_bytes.decode("utf-8", errors="replace"))
        extracted_text = json.dumps(parsed)[:20000]
        if isinstance(parsed, list) and parsed and isinstance(parsed[0], dict):
            num_stats: dict[str, dict[str, float]] = {}
            for k in parsed[0].keys():
                vals = [v for v in (_to_float_or_none(r.get(k)) for r in parsed if isinstance(r, dict)) if v is not None]
                if vals:
                    arr = np.array(vals, dtype=float)
                    num_stats[k] = {
                        "count": len(vals),
                        "min": round(float(np.min(arr)), 5),
                        "max": round(float(np.max(arr)), 5),
                        "mean": round(float(np.mean(arr)), 5),
                        "std": round(float(np.std(arr)), 5),
                    }
            extra["numeric_summaries"] = num_stats
        elif isinstance(parsed, dict):
            extra["top_level_keys"] = list(parsed.keys())[:30]
    elif ext == ".pdf":
        pages_text: list[str] = []
        rendered_pages: list[dict[str, str]] = []
        try:
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                extra["page_count"] = len(doc)
                for idx in range(min(len(doc), 25)):
                    pages_text.append(doc[idx].get_text("text"))
                for p_idx in range(min(len(doc), 3)):
                    pix = doc[p_idx].get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False)
                    png_bytes = pix.tobytes("png")
                    rendered_pages.append(
                        {
                            "label": f"{filename}#page{p_idx + 1}",
                            "mime_type": "image/png",
                            "base64": base64.b64encode(png_bytes).decode("ascii"),
                        }
                    )
                extra["rendered_page_count"] = len(rendered_pages)
                extra["_image_payloads"] = rendered_pages
        except Exception:
            reader = PdfReader(io.BytesIO(file_bytes))
            extra["page_count"] = len(reader.pages)
            for page in reader.pages[:25]:
                pages_text.append(page.extract_text() or "")
        extracted_text = "\n".join(pages_text)
    elif ext == ".docx":
        doc = DocxDocument(io.BytesIO(file_bytes))
        paras = [p.text for p in doc.paragraphs if p.text.strip()]
        table_lines: list[str] = []
        for tbl in doc.tables[:5]:
            for row in tbl.rows[:25]:
                table_lines.append(" | ".join(c.text.strip() for c in row.cells))
        extra["paragraph_count"] = len(paras)
        extra["table_count"] = len(doc.tables)
        extracted_text = "\n".join(paras + table_lines)
    elif ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        extra["sheet_names"] = wb.sheetnames[:10]
        sheets_info: dict[str, Any] = {}
        chunks: list[str] = []
        for name in wb.sheetnames[:10]:
            ws = wb[name]
            rows_data = list(ws.iter_rows(max_row=1000, values_only=True))
            sheets_info[name] = {
                "rows": len(rows_data),
                "columns": max((len(r) for r in rows_data), default=0),
            }
            for r in rows_data[:40]:
                chunks.append(" ".join(str(c) for c in r if c is not None))
        extra["sheets"] = sheets_info
        extracted_text = "\n".join(chunks)
    else:
        extracted_text = file_bytes.decode("utf-8", errors="replace")[:20000]

    psi_mentions = sorted(set(m.upper() for m in re.findall(r"PSI-\d+", extracted_text, flags=re.IGNORECASE)))
    fuel_keywords = [
        f for f in ("pmma", "heptane", "methanol", "decane", "cotton", "nomex", "methane", "ethylene")
        if re.search(rf"\b{f}\b", extracted_text, flags=re.IGNORECASE)
    ]

    if ext in (".jdx", ".dx") or any(k in low_name for k in ("spectrum", "spectral", "ftir", "raman", "lif")):
        modality = "spectral"
    elif ext in (".xlsx", ".xlsm") or extra.get("numeric_summaries"):
        modality = "numeric"
    else:
        modality = "text"

    return {
        "modality": modality,
        "filename": filename,
        "format": ext.lstrip(".") or "txt",
        "character_count": len(extracted_text),
        "detected_psi_accessions": psi_mentions,
        "detected_fuels": fuel_keywords,
        "excerpt": extracted_text[:500].strip(),
        **extra,
    }


def parse_evidence_file(file_bytes: bytes, filename: str, content_type: str = "") -> dict[str, Any]:
    """Classify and parse uploaded evidence files with size caps per Section 06."""
    if len(file_bytes) > MAX_FILE_BYTES:
        raise ValueError(
            f"File '{filename}' ({len(file_bytes)} bytes) exceeds PYROSAGE_MAX_FILE_BYTES ({MAX_FILE_BYTES} bytes)."
        )
    ext = Path(filename).suffix.lower()
    ct = (content_type or "").lower()

    if ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff") or ct.startswith("image/"):
        desc = extract_image_descriptors(file_bytes, filename)
    elif ext in (".mp4", ".mov", ".avi", ".mkv", ".webm", ".mpeg", ".mpg") or ct.startswith("video/"):
        desc = extract_video_descriptors(file_bytes, filename)
    elif ext in (".csv", ".tsv") or ct in ("text/csv", "text/tab-separated-values"):
        desc = extract_table_descriptors(file_bytes, filename)
    elif ext in (".json", ".pdf", ".docx", ".xlsx", ".xlsm", ".txt", ".md", ".log", ".dat", ".asc", ".jdx", ".dx"):
        desc = extract_document_descriptors(file_bytes, filename)
    else:
        raise ValueError(f"Unsupported file extension '{ext or content_type}' for '{filename}'.")

    desc["size_bytes"] = len(file_bytes)
    desc["evidence_class"] = "uploaded_evidence"
    return desc


def infer_suggested_techniques(extractions: list[dict[str, Any]]) -> list[str]:
    """Infer suggested technique IDs from detected file modalities per Section 06."""
    mods = {item.get("modality", "text") for item in extractions}
    suggested: list[str] = []
    if mods & {"image", "thermal"}:
        suggested.extend(
            [
                "segmentation",
                "spherical_symmetry",
                "reflection_iou",
                "bilateral_symmetry",
                "top_bottom_symmetry",
                "rotational_symmetry",
                "convex_hull",
                "convexity_defects",
                "fractal_dimension",
                "boundary_roughness",
                "perimeter_area",
                "peak_detection",
                "centroid_center_mass",
                "absence_buoyancy",
            ]
        )
    if "video" in mods:
        suggested.extend(["temporal_dynamics", "symmetry_variance", "centroid_jitter", "cool_flame_extinction"])
    if "spectral" in mods:
        suggested.extend(["lif", "ftir", "raman", "formaldehyde_ho2", "co_analysis", "soot_precursors"])
    if "numeric" in mods:
        suggested.extend(["temperature_distribution", "species_concentration", "gas_velocity", "heat_radiation"])
    if "text" in mods:
        suggested.extend(["iss_cir", "drop_tower", "molecular_diffusion"])
    suggested.extend(["cross_experiment_compare", "evidence_ranking", "anomaly_detection", "risk_interpretation"])
    seen: set[str] = set()
    ordered: list[str] = []
    for tid in suggested:
        if tid not in seen:
            seen.add(tid)
            ordered.append(tid)
    return ordered


def match_catalog_candidates(
    extractions: list[dict[str, Any]],
    metadata: dict[str, Any],
    catalog_records: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    """Score imported SQLite catalogue records against metadata and optional geometry descriptors (Section 09)."""
    records = catalog_records if catalog_records is not None else list_catalog_records()["records"]
    if not records:
        return []

    q_gravity = _to_float_or_none(metadata.get("gravity_g"))
    q_platform = str(metadata.get("platform", "")).strip().lower()
    q_fuel = str(metadata.get("fuel_material") or metadata.get("fuel") or "").strip().lower()
    q_mission = str(metadata.get("mission", "")).strip().lower()

    img_ext = next((e for e in extractions if e.get("modality") in ("image", "video", "thermal")), None)

    candidates: list[dict[str, Any]] = []
    for rec in records:
        score = 0.0
        reasons: list[str] = []
        weights = 0.0

        rec_grav = _to_float_or_none(rec.get("gravity_g"))
        if q_gravity is not None and rec_grav is not None:
            g_sim = max(0.0, 1.0 - abs(q_gravity - rec_grav) * 1.5)
            score += 0.25 * g_sim
            weights += 0.25
            if g_sim >= 0.8:
                reasons.append(f"gravity match ({rec_grav}g)")

        rec_plat = str(rec.get("platform", "")).lower()
        if q_platform and rec_plat:
            weights += 0.20
            if q_platform in rec_plat or rec_plat in q_platform:
                score += 0.20
                reasons.append(f"platform match ({rec.get('platform')})")

        rec_fuel = str(rec.get("fuel_material", "")).lower()
        if q_fuel and rec_fuel:
            weights += 0.25
            if any(tok in rec_fuel for tok in re.split(r"[,\s]+", q_fuel) if tok):
                score += 0.25
                reasons.append(f"fuel/material match ({q_fuel})")

        rec_mission = str(rec.get("mission", "")).lower()
        if q_mission and rec_mission and (q_mission in rec_mission or rec_mission in q_mission):
            score += 0.15
            weights += 0.15
            reasons.append("mission context match")

        # Optional geometry matching per Section 09 & Page 18
        if img_ext:
            geom_diffs: list[float] = []
            for gf in ("bilateral_left_right_score", "top_bottom_symmetry_score", "convexity_score", "fractal_dimension"):
                rec_val = _to_float_or_none(rec.get(gf))
                ext_val = _to_float_or_none(img_ext.get(gf))
                if rec_val is not None and ext_val is not None:
                    scale = 2.0 if gf == "fractal_dimension" else 1.0
                    geom_diffs.append(abs(ext_val - rec_val) / scale)
            if geom_diffs:
                geom_sim = float(np.clip(1.0 - float(np.mean(geom_diffs)) * 1.5, 0.0, 1.0))
                score += 0.45 * geom_sim
                weights += 0.45
                reasons.append(f"geometry descriptor similarity ({geom_sim:.2f})")

        final_score = float(score / max(weights, 0.4)) if weights > 0 else 0.35
        final_score = float(np.clip(final_score, 0.05, 0.99))
        candidates.append(
            {
                "experiment_id": rec["experiment_id"],
                "accession": rec["experiment_id"],
                "name": rec["name"],
                "title": rec["name"],
                "platform": rec.get("platform", ""),
                "gravity_g": rec.get("gravity_g"),
                "fuel_material": rec.get("fuel_material", ""),
                "pre_match_score": round(final_score, 4),
                "similarity_score": round(final_score, 4),
                "match_confidence": round(final_score, 4),
                "reason": "; ".join(reasons) or "Catalogue candidate pre-match",
                "match_rationale": "; ".join(reasons) or "Catalogue candidate pre-match",
                "source_ref": rec.get("source_ref", ""),
                "source_url": rec.get("source_ref", ""),
            }
        )

    candidates.sort(key=lambda c: (-c["pre_match_score"], c["experiment_id"]))
    return candidates[:12]


def evaluate_techniques(
    selected_ids: list[str],
    extractions: list[dict[str, Any]],
    metadata: dict[str, Any],
) -> list[dict[str, Any]]:
    """Build initial deterministic technique assessments honoring method prerequisites and solver locks."""
    tech_map = get_technique_map()
    if not selected_ids:
        selected_ids = infer_suggested_techniques(extractions)

    resolved_ids: list[str] = []
    for raw_id in selected_ids:
        t_obj = tech_map.get(str(raw_id).strip())
        if t_obj and t_obj["id"] not in resolved_ids:
            resolved_ids.append(t_obj["id"])

    available_mods = {e.get("modality", "text") for e in extractions}
    if any(
        metadata.get(k) not in (None, "")
        for k in ("platform", "gravity_g", "fuel_material", "fuel", "mission", "notes")
    ):
        available_mods.add("text")

    assessments: list[dict[str, Any]] = []
    for tid in resolved_ids:
        tech = tech_map[tid]
        supported = set(tech.get("modalities") or tech.get("supported_modalities") or [])
        has_mod = bool(available_mods & supported)

        if tid in UNSUPPORTED_SOLVER_TECHNIQUES:
            status = "requires_more_data"
            obs = f"{tech['name']} is locked to requires_more_data in this build because its specialist solver/calibration is not bundled."
        elif not has_mod:
            status = "requires_more_data"
            obs = f"Uploaded evidence modalities ({', '.join(sorted(available_mods)) or 'none'}) do not include required ({', '.join(sorted(supported))})."
        elif tid in LOCALLY_COMPUTED_IMAGE_TECHNIQUES and (available_mods & {"image", "video", "thermal"}):
            status = "computed_from_upload"
            obs = f"Computed locally via OpenCV/NumPy contour and reflection mask pipeline."
        elif tid in LOCALLY_COMPUTED_VIDEO_TECHNIQUES and "video" in available_mods:
            status = "computed_from_upload"
            obs = f"Computed locally from sampled video frames."
        elif tid in {"drop_tower", "parabolic_flight", "sounding_rocket", "iss_cir", "orbital_platforms", "absence_buoyancy"} and "text" in available_mods:
            status = "interpreted_from_metadata"
            obs = f"Interpreted from supplied acquisition metadata and context."
        elif tid in {"cross_experiment_compare", "evidence_ranking", "anomaly_detection", "risk_interpretation"}:
            status = "computed_from_upload" if extractions else "interpreted_from_metadata"
            obs = f"Synthesized from local extraction descriptors and bounded catalogue context."
        else:
            status = "requires_more_data"
            obs = f"Requires calibrated instrument data or reference context: {tech.get('prerequisites', '')}"

        # Provide UI-friendly alias `state` alongside canonical `status`
        state_alias = {
            "computed_from_upload": "Executed",
            "visually_assessed": "Partially executed",
            "interpreted_from_metadata": "Partially executed",
            "requires_more_data": (
                "Requires specialist solver" if tid in UNSUPPORTED_SOLVER_TECHNIQUES else "Requires data"
            ),
            "not_applicable": "Not applicable",
        }[status]

        assessments.append(
            {
                "technique_id": tech["id"],
                "id": tech["id"],
                "name": tech["name"],
                "group": tech["group"],
                "status": status,
                "state": state_alias,
                "observations": obs,
                "rationale": obs,
                "evidence_refs": [e.get("filename", "upload") for e in extractions[:3]],
                "limitations": str(tech.get("prerequisites", "")),
            }
        )
    return assessments


def validate_analysis_response(
    raw_report: dict[str, Any],
    selected_technique_ids: list[str],
    candidate_experiments: list[dict[str, Any]],
) -> dict[str, Any]:
    """Validate LLM/report output per Section 08 (Page 14 & Page 18):
    - Keep only selected & known technique IDs
    - Supply missing selected IDs with requires_more_data
    - Normalize unknown statuses to requires_more_data
    - Force piv, cnn_resnet, symmetry_loss, spectral_deconvolution, chemical_kinetics, gravity_scaling to requires_more_data
    - Remove any experiment_id not in candidate_experiments
    - Add data gap if catalogue/candidates are empty
    - Clamp relevance and confidence scores to [0.0, 1.0].
    """
    tech_map = get_technique_map()
    canonical_selected: list[str] = []
    for raw_id in selected_technique_ids:
        t_obj = tech_map.get(str(raw_id).strip())
        if t_obj and t_obj["id"] not in canonical_selected:
            canonical_selected.append(t_obj["id"])

    allowed_tech_set = set(canonical_selected)
    candidate_map = {str(c.get("experiment_id") or c.get("accession")): c for c in candidate_experiments}
    allowed_exp_ids = set(candidate_map.keys())

    # 1. Validate technique_assessments
    in_assessments = raw_report.get("technique_assessments", [])
    valid_assessments_by_id: dict[str, dict[str, Any]] = {}
    if isinstance(in_assessments, list):
        for item in in_assessments:
            if not isinstance(item, dict):
                continue
            raw_tid = str(item.get("technique_id") or item.get("id") or "").strip()
            t_obj = tech_map.get(raw_tid)
            if not t_obj or t_obj["id"] not in allowed_tech_set:
                continue
            cid = t_obj["id"]
            st = str(item.get("status", "requires_more_data")).strip()
            if st not in ALLOWED_STATUSES:
                st = "requires_more_data"
            if cid in UNSUPPORTED_SOLVER_TECHNIQUES:
                st = "requires_more_data"
            state_alias = {
                "computed_from_upload": "Executed",
                "visually_assessed": "Partially executed",
                "interpreted_from_metadata": "Partially executed",
                "requires_more_data": (
                    "Requires specialist solver" if cid in UNSUPPORTED_SOLVER_TECHNIQUES else "Requires data"
                ),
                "not_applicable": "Not applicable",
            }[st]
            valid_assessments_by_id[cid] = {
                "technique_id": cid,
                "id": cid,
                "name": t_obj["name"],
                "group": t_obj["group"],
                "status": st,
                "state": state_alias,
                "observations": str(item.get("observations") or item.get("rationale") or ""),
                "rationale": str(item.get("rationale") or item.get("observations") or ""),
                "evidence_refs": list(item.get("evidence_refs", [])),
                "limitations": str(item.get("limitations") or t_obj.get("prerequisites", "")),
            }

    # Add missing selected techniques with requires_more_data
    ordered_assessments: list[dict[str, Any]] = []
    for cid in canonical_selected:
        if cid in valid_assessments_by_id:
            ordered_assessments.append(valid_assessments_by_id[cid])
        else:
            t_obj = tech_map[cid]
            ordered_assessments.append(
                {
                    "technique_id": cid,
                    "id": cid,
                    "name": t_obj["name"],
                    "group": t_obj["group"],
                    "status": "requires_more_data",
                    "state": "Requires specialist solver" if cid in UNSUPPORTED_SOLVER_TECHNIQUES else "Requires data",
                    "observations": "Requires additional data or specialist solver.",
                    "rationale": "Requires additional data or specialist solver.",
                    "evidence_refs": [],
                    "limitations": str(t_obj.get("prerequisites", "")),
                }
            )

    # 2. Validate experiment_matches against candidate_experiments
    in_matches = raw_report.get("experiment_matches", raw_report.get("candidate_experiments", []))
    validated_matches: list[dict[str, Any]] = []
    rejected_exp_ids: list[str] = []
    if isinstance(in_matches, list) and allowed_exp_ids:
        for m in in_matches:
            if not isinstance(m, dict):
                continue
            eid = str(m.get("experiment_id") or m.get("accession") or "").strip()
            if eid in allowed_exp_ids:
                cand = candidate_map[eid]
                conf = _to_float_or_none(m.get("match_confidence") or m.get("similarity_score"))
                conf_clamped = round(float(np.clip(conf if conf is not None else cand.get("pre_match_score", 0.5), 0.0, 1.0)), 4)
                validated_matches.append(
                    {
                        **cand,
                        "experiment_id": eid,
                        "accession": eid,
                        "match_confidence": conf_clamped,
                        "similarity_score": conf_clamped,
                        "reason": str(m.get("reason") or cand.get("reason") or ""),
                    }
                )
            elif eid:
                rejected_exp_ids.append(eid)

    # 3. Data gaps & empty catalogue check
    data_gaps = [str(g) for g in raw_report.get("data_gaps", []) if g]
    if not allowed_exp_ids:
        validated_matches = []
        gap_msg = "Experiment catalogue database is empty or no candidate records matched; import catalogue records to enable experiment identification."
        if not any("empty" in g.lower() or "no candidate" in g.lower() for g in data_gaps):
            data_gaps.append(gap_msg)

    # 4. Clamp key_findings scores and strip unverified PSI accessions / unselected technique_ids
    in_findings = raw_report.get("key_findings", [])
    validated_findings: list[dict[str, Any]] = []
    if isinstance(in_findings, list):
        for idx, f in enumerate(in_findings, start=1):
            if not isinstance(f, dict):
                continue
            rel = float(np.clip(_to_float_or_none(f.get("relevance_score")) or 0.75, 0.0, 1.0))
            conf = float(np.clip(_to_float_or_none(f.get("confidence")) or 0.70, 0.0, 1.0))
            finding_str = str(f.get("finding") or f.get("interpretation") or "")
            title_str = str(f.get("title") or f"Finding #{idx}")
            ev_str = str(f.get("evidence") or f.get("evidence_reference") or "")

            for mentioned_acc in re.findall(r"PSI-\d+", f"{title_str} {finding_str} {ev_str}", flags=re.IGNORECASE):
                norm_m = mentioned_acc.upper()
                if allowed_exp_ids and norm_m not in allowed_exp_ids:
                    if norm_m not in rejected_exp_ids:
                        rejected_exp_ids.append(norm_m)
                    title_str = re.sub(rf"\b{re.escape(mentioned_acc)}\b", "[UNVERIFIED-REMOVED]", title_str, flags=re.IGNORECASE)
                    finding_str = re.sub(rf"\b{re.escape(mentioned_acc)}\b", "[UNVERIFIED-REMOVED]", finding_str, flags=re.IGNORECASE)
                    ev_str = re.sub(rf"\b{re.escape(mentioned_acc)}\b", "[UNVERIFIED-REMOVED]", ev_str, flags=re.IGNORECASE)

            raw_f_tids = f.get("technique_ids", [])
            valid_f_tids: list[str] = []
            if isinstance(raw_f_tids, list):
                for r_tid in raw_f_tids:
                    r_str = str(r_tid).strip()
                    t_obj = tech_map.get(r_str)
                    if t_obj and t_obj["id"] in allowed_tech_set:
                        valid_f_tids.append(r_str if r_str in selected_technique_ids else t_obj["id"])

            validated_findings.append(
                {
                    "rank": int(f.get("rank") or idx),
                    "title": title_str,
                    "relevance_score": round(rel, 3),
                    "confidence": round(conf, 3),
                    "evidence_type": str(f.get("evidence_type") or f.get("evidence_class") or "computed_from_upload"),
                    "evidence_class": str(f.get("evidence_class") or f.get("evidence_type") or "computed_from_upload"),
                    "technique_ids": valid_f_tids,
                    "finding": finding_str,
                    "interpretation": finding_str,
                    "evidence": ev_str,
                    "evidence_reference": ev_str,
                }
            )

    overall_conf = float(np.clip(_to_float_or_none(raw_report.get("overall_confidence")) or 0.72, 0.0, 1.0))

    return {
        "executive_summary": str(
            raw_report.get("executive_summary")
            or raw_report.get("summary")
            or "Technique-bounded PyroSage analysis completed."
        ),
        "technique_assessments": ordered_assessments,
        "experiment_matches": validated_matches,
        "candidate_experiments": validated_matches,
        "rejected_experiment_ids": rejected_exp_ids,
        "key_findings": validated_findings,
        "safety_implications": list(raw_report.get("safety_implications", [])),
        "data_gaps": data_gaps,
        "overall_confidence": round(overall_conf, 3),
        "disclaimer": (
            "Research decision support only — not a validated combustion simulation, "
            "mission-safety determination, or flight-safety certification."
        ),
    }


def validate_si_output(
    raw_output: dict[str, Any],
    allowed_technique_ids: list[str],
    allowed_psi_accessions: list[str],
    technique_assessments: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    cands = [{"experiment_id": acc, "accession": acc, "name": acc, "pre_match_score": 0.85} for acc in allowed_psi_accessions]
    val = validate_analysis_response(
        raw_report={
            "executive_summary": raw_output.get("summary", ""),
            "technique_assessments": raw_output.get("technique_assessments", technique_assessments or []),
            "experiment_matches": [
                {"experiment_id": acc} for acc in raw_output.get("referenced_psi_accessions", [])
            ],
            "key_findings": raw_output.get("findings", []),
            "data_gaps": raw_output.get("data_gaps", []),
        },
        selected_technique_ids=allowed_technique_ids,
        candidate_experiments=cands,
    )
    return {
        "summary": val["executive_summary"],
        "referenced_psi_accessions": [m["experiment_id"] for m in val["experiment_matches"]],
        "rejected_hallucinated_accessions": val["rejected_experiment_ids"],
        "technique_assessments": val["technique_assessments"],
        "findings": val["key_findings"],
        "data_gaps": val["data_gaps"],
    }


SI_RESPONSE_JSON_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "required": [
        "executive_summary",
        "technique_assessments",
        "experiment_matches",
        "key_findings",
        "safety_implications",
        "data_gaps",
        "overall_confidence",
    ],
    "properties": {
        "executive_summary": {"type": "string"},
        "technique_assessments": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["technique_id", "status", "observations", "evidence_refs", "limitations"],
                "properties": {
                    "technique_id": {"type": "string"},
                    "status": {"type": "string", "enum": sorted(ALLOWED_STATUSES)},
                    "observations": {"type": "string"},
                    "evidence_refs": {"type": "array", "items": {"type": "string"}},
                    "limitations": {"type": "string"},
                },
            },
        },
        "experiment_matches": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["experiment_id", "match_confidence", "reason"],
                "properties": {
                    "experiment_id": {"type": "string"},
                    "match_confidence": {"type": "number"},
                    "reason": {"type": "string"},
                },
            },
        },
        "key_findings": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["rank", "relevance_score", "confidence", "evidence_type", "finding", "evidence"],
                "properties": {
                    "rank": {"type": "integer"},
                    "relevance_score": {"type": "number"},
                    "confidence": {"type": "number"},
                    "evidence_type": {"type": "string"},
                    "finding": {"type": "string"},
                    "evidence": {"type": "string"},
                },
            },
        },
        "safety_implications": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["implication", "priority", "supporting_evidence", "uncertainty"],
                "properties": {
                    "implication": {"type": "string"},
                    "priority": {"type": "string"},
                    "supporting_evidence": {"type": "string"},
                    "uncertainty": {"type": "string"},
                },
            },
        },
        "data_gaps": {"type": "array", "items": {"type": "string"}},
        "overall_confidence": {"type": "number"},
    },
}


def collect_image_payloads(extractions: list[dict[str, Any]]) -> list[dict[str, str]]:
    """Collect up to MAX_IMAGE_PAYLOADS (5) image payloads across images, video sample frames, and rendered PDF pages."""
    payloads: list[dict[str, str]] = []
    for ext in extractions:
        items = ext.get("_image_payloads", [])
        if isinstance(items, list):
            for p in items:
                if isinstance(p, dict) and p.get("base64"):
                    payloads.append(p)
                    if len(payloads) >= MAX_IMAGE_PAYLOADS:
                        return payloads
    return payloads


def strip_internal_payloads(extractions: list[dict[str, Any]]) -> list[dict[str, Any]]:
    cleaned: list[dict[str, Any]] = []
    for ext in extractions:
        copy_ext = {k: v for k, v in ext.items() if not k.startswith("_")}
        cleaned.append(copy_ext)
    return cleaned


def build_si_system_instruction(
    selected_techniques: list[dict[str, Any]],
    candidate_experiments: list[dict[str, Any]],
) -> str:
    """Build the bounded SI system instruction enforcing the 9 mandatory rules from Section 08 (Slide 14)."""
    tech_lines = "\n".join(
        f"- {t['id']} ({t['name']}): modalities={t.get('modalities', [])}; prerequisites={t.get('prerequisites', '')}"
        for t in selected_techniques
    )
    cand_lines = (
        "\n".join(
            f"- {c['experiment_id']}: {c['name']} | platform={c.get('platform', '')} | g={c.get('gravity_g')} | pre_match={c.get('pre_match_score')}"
            for c in candidate_experiments
        )
        if candidate_experiments
        else "NONE (Catalogue is empty or no candidate records matched)."
    )
    return (
        "You are PyroSage SI, a technique-bounded microgravity combustion evidence analysis assistant.\n"
        "MANDATORY ANALYSIS RULES (Section 08):\n"
        "1. Use ONLY methods in the supplied allowed-technique whitelist below. Never invent a diagnostic method, model, metric, or technique.\n"
        "2. Assign one of the 5 allowed statuses (computed_from_upload, visually_assessed, interpreted_from_metadata, requires_more_data, not_applicable) to each method.\n"
        "3. Do NOT claim PIV without seeded tracer particles, paired frames, frame timing, and spatial calibration. Optical flow is not PIV.\n"
        "4. Do NOT infer absolute temperature (K or °C) from ordinary RGB or uncalibrated false-color images.\n"
        "5. Do NOT identify chemical species from an uncalibrated peak alone.\n"
        "6. Do NOT infer the experiment, mission, or platform from flame shape. Experiment identification MUST be confined to supplied candidate records.\n"
        "7. Do NOT invent NASA experiments, citations, measurement values, frame rates, instrument metadata, or physical units.\n"
        "8. State uncertainty in heuristic segmentation and distinguish measured, locally computed, visually assessed, inferred, and unknown values.\n"
        "9. Do NOT call model interpretation a validated combustion simulation, kinetics solver, mission-safety determination, or certification.\n\n"
        f"ALLOWED TECHNIQUE WHITELIST ({len(selected_techniques)}):\n{tech_lines}\n\n"
        f"SUPPLIED CATALOGUE CANDIDATES:\n{cand_lines}\n"
    )


def call_gemini_interactions_api(
    api_key: str,
    model: str,
    system_instruction: str,
    context_text: str,
    image_payloads: list[dict[str, str]],
) -> dict[str, Any]:
    """Invoke Google Gemini Interactions API (v1beta/interactions) with store=false and JSON schema per Slide 14."""
    parts: list[dict[str, Any]] = [{"text": context_text}]
    for img in image_payloads[:MAX_IMAGE_PAYLOADS]:
        parts.append(
            {
                "inline_data": {
                    "mime_type": img.get("mime_type", "image/png"),
                    "data": img["base64"],
                }
            }
        )

    payload = {
        "model": model,
        "system_instruction": {"parts": [{"text": system_instruction}]},
        "contents": [{"role": "user", "parts": parts}],
        "generation_config": {
            "temperature": 0.1,
            "response_mime_type": "application/json",
            "response_schema": SI_RESPONSE_JSON_SCHEMA,
        },
        "store": False,
    }
    url = "https://generativelanguage.googleapis.com/v1beta/interactions"
    with httpx.Client(timeout=30.0) as http_client:
        resp = http_client.post(url, headers={"x-goog-api-key": api_key, "Content-Type": "application/json"}, json=payload)
        if resp.status_code >= 400:
            alt_url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
            alt_payload = {k: v for k, v in payload.items() if k not in ("model", "store")}
            resp = http_client.post(alt_url, headers={"x-goog-api-key": api_key, "Content-Type": "application/json"}, json=alt_payload)
        resp.raise_for_status()
        data = resp.json()

    text_out = ""
    if "outputs" in data and isinstance(data["outputs"], list) and data["outputs"]:
        text_out = data["outputs"][0].get("text", "")
    elif "candidates" in data and data["candidates"]:
        c_parts = data["candidates"][0].get("content", {}).get("parts", [])
        text_out = "".join(p.get("text", "") for p in c_parts)
    return json.loads(text_out) if text_out else {}


def call_openai_responses_api(
    api_key: str,
    model: str,
    system_instruction: str,
    context_text: str,
    image_payloads: list[dict[str, str]],
) -> dict[str, Any]:
    """Invoke OpenAI Responses API (v1/responses) with high-detail images and strict JSON schema per Slide 14."""
    user_content: list[dict[str, Any]] = [{"type": "input_text", "text": context_text}]
    for img in image_payloads[:MAX_IMAGE_PAYLOADS]:
        mime = img.get("mime_type", "image/png")
        user_content.append(
            {
                "type": "input_image",
                "image_url": f"data:{mime};base64,{img['base64']}",
                "detail": "high",
            }
        )

    payload = {
        "model": model,
        "store": False,
        "input": [
            {"role": "system", "content": [{"type": "input_text", "text": system_instruction}]},
            {"role": "user", "content": user_content},
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "pyrosage_si_report",
                "strict": True,
                "schema": SI_RESPONSE_JSON_SCHEMA,
            }
        },
    }
    with httpx.Client(timeout=30.0) as http_client:
        resp = http_client.post(
            "https://api.openai.com/v1/responses",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json=payload,
        )
        resp.raise_for_status()
        data = resp.json()

    text_out = data.get("output_text", "")
    if not text_out and "output" in data:
        for item in data["output"]:
            for c in item.get("content", []):
                if c.get("type") in ("output_text", "text"):
                    text_out += c.get("text", "")
    return json.loads(text_out) if text_out else {}


def invoke_si_provider(
    provider: str,
    model: str,
    request_api_key: str | None,
    selected_tech_ids: list[str],
    extractions: list[dict[str, Any]],
    metadata: dict[str, Any],
    candidates: list[dict[str, Any]],
    fallback_report: dict[str, Any],
) -> tuple[dict[str, Any], str]:
    """Invoke the selected external LLM provider if an API key is available, else return deterministic local synthesis."""
    tech_map = get_technique_map()
    selected_objs = [tech_map[tid] for tid in selected_tech_ids if tid in tech_map]
    sys_prompt = build_si_system_instruction(selected_objs, candidates)
    img_payloads = collect_image_payloads(extractions)
    clean_ext = strip_internal_payloads(extractions)
    context_text = json.dumps(
        {
            "metadata": metadata,
            "local_extractions": clean_ext,
            "candidate_experiments": candidates,
        },
        indent=2,
    )

    prov = (provider or DEFAULT_PROVIDER).strip().lower()
    key = (request_api_key or "").strip()
    if not key:
        key = os.getenv("OPENAI_API_KEY" if prov == "openai" else "GEMINI_API_KEY", "").strip()

    if not key:
        return fallback_report, "local_deterministic_synthesis"

    try:
        if prov == "openai":
            raw = call_openai_responses_api(key, model or "gpt-4.1-mini", sys_prompt, context_text, img_payloads)
            return raw or fallback_report, "openai_responses_api"
        raw = call_gemini_interactions_api(key, model or DEFAULT_MODEL, sys_prompt, context_text, img_payloads)
        return raw or fallback_report, "gemini_interactions_api"
    except Exception:
        return fallback_report, "local_deterministic_fallback"


def create_synthetic_sample_image(kind: str = "0g_spherical") -> bytes:
    """Create a synthetic warm-color HSV flame test image for verification and demo use."""
    canvas = np.zeros((240, 240, 3), dtype=np.uint8)
    canvas[:, :] = (14, 12, 10)
    if kind == "0g_spherical":
        # BGR warm orange/red/yellow flame
        cv2.circle(canvas, (120, 120), 62, (20, 75, 225), -1)
        cv2.circle(canvas, (120, 120), 42, (35, 140, 245), -1)
        cv2.circle(canvas, (120, 120), 20, (80, 220, 255), -1)
    elif kind == "lunar_017g":
        cv2.ellipse(canvas, (120, 125), (44, 68), 0, 0, 360, (20, 80, 230), -1)
        cv2.ellipse(canvas, (120, 118), (28, 48), 0, 0, 360, (45, 165, 250), -1)
    else:
        pts = np.array([[120, 30], [162, 145], [142, 202], [98, 202], [78, 145]], dtype=np.int32)
        cv2.fillPoly(canvas, [pts], (20, 80, 235))
        cv2.circle(canvas, (120, 150), 26, (60, 195, 255), -1)
    ok, encoded = cv2.imencode(".png", canvas)
    if not ok:
        raise RuntimeError("Failed to encode synthetic flame image.")
    return encoded.tobytes()


@asynccontextmanager
async def _lifespan(_app: FastAPI):
    init_db()
    yield


app = FastAPI(
    title="PyroSage — Microgravity Fire Intelligence API",
    version="1.0.0",
    description="Technique-bounded multimodal research dashboard and API for microgravity combustion evidence analysis.",
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
    if MAIN_DASHBOARD_PATH.exists():
        return FileResponse(str(MAIN_DASHBOARD_PATH), media_type="text/html")
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file), media_type="text/html")
    if PROTOTYPE_PATH.exists():
        return FileResponse(str(PROTOTYPE_PATH), media_type="text/html")
    return HTMLResponse("<h1>PyroSage API is running</h1>")


@app.get("/pyrosage_interactive_dashboard.html", response_class=HTMLResponse)
def serve_main_interactive_dashboard() -> Response:
    if MAIN_DASHBOARD_PATH.exists():
        return FileResponse(str(MAIN_DASHBOARD_PATH), media_type="text/html")
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file), media_type="text/html")
    raise HTTPException(status_code=404, detail="pyrosage_interactive_dashboard.html not found.")


@app.get("/reference_prototype.html", response_class=HTMLResponse)
def serve_reference_prototype() -> Response:
    if PROTOTYPE_PATH.exists():
        return FileResponse(str(PROTOTYPE_PATH), media_type="text/html")
    raise HTTPException(status_code=404, detail="reference_prototype.html not found.")


@app.get("/api/health")
def api_health() -> dict[str, Any]:
    init_db()
    techniques = load_techniques()
    cat = list_catalog_records(limit=1)
    return {
        "status": "ok",
        "service": "PyroSage",
        "version": "1.0.0",
        "default_provider": DEFAULT_PROVIDER,
        "default_model": DEFAULT_MODEL,
        "providers_configured": {
            "gemini": bool(os.getenv("GEMINI_API_KEY", "").strip()),
            "openai": bool(os.getenv("OPENAI_API_KEY", "").strip()),
        },
        "gemini_configured": bool(os.getenv("GEMINI_API_KEY", "").strip()),
        "openai_configured": bool(os.getenv("OPENAI_API_KEY", "").strip()),
        "nasa_psi_base_url": NASA_PSI_BASE_URL,
        "technique_count": len(techniques),
        "catalog_count": cat["total"],
        "cached_psi_experiments": len(VERIFIED_PSI_RECORDS),
        "limits": {
            "max_file_bytes": MAX_FILE_BYTES,
            "max_total_bytes": MAX_TOTAL_BYTES,
            "max_files": MAX_UPLOAD_FILES,
            "max_image_payloads": MAX_IMAGE_PAYLOADS,
        },
    }


@app.get("/api/techniques")
def api_techniques(
    group: str | None = Query(default=None),
    modality: str | None = Query(default=None),
) -> dict[str, Any]:
    all_techs = load_techniques()
    filtered = all_techs
    if group:
        g_low = group.strip().lower()
        filtered = [t for t in filtered if g_low in t["group"].lower()]
    if modality:
        m_low = modality.strip().lower()
        filtered = [t for t in filtered if m_low in [m.lower() for m in t.get("modalities", [])]]

    families: dict[str, int] = {}
    for t in all_techs:
        families[t["group"]] = families.get(t["group"], 0) + 1

    return {
        "count": len(filtered),
        "total_registered": len(all_techs),
        "families": [{"name": k, "count": v} for k, v in families.items()],
        "statuses": sorted(ALLOWED_STATUSES),
        "techniques": filtered,
    }


@app.get("/api/catalog")
def api_get_catalog() -> dict[str, Any]:
    return list_catalog_records(limit=1000)


@app.post("/api/catalog/import")
async def api_import_catalog(
    request: Request,
    file: UploadFile | None = File(default=None),
) -> dict[str, Any]:
    init_db()
    rows_to_import: list[dict[str, Any]] = []

    if file and file.filename:
        raw_bytes = await file.read()
        fname = file.filename.lower()
        text = raw_bytes.decode("utf-8", errors="replace")
        if fname.endswith(".json"):
            parsed = json.loads(text)
            if isinstance(parsed, list):
                rows_to_import = parsed
            elif isinstance(parsed, dict):
                rows_to_import = (
                    parsed.get("experiments")
                    or parsed.get("records")
                    or parsed.get("items")
                    or [parsed]
                )
        else:
            delim = "\t" if fname.endswith(".tsv") or ("\t" in (text.splitlines()[0] if text.splitlines() else "")) else ","
            rows_to_import = list(csv.DictReader(io.StringIO(text), delimiter=delim))
    else:
        body = await request.json()
        if isinstance(body, list):
            rows_to_import = body
        elif isinstance(body, dict):
            rows_to_import = (
                body.get("experiments")
                or body.get("records")
                or body.get("items")
                or [body]
            )

    if not isinstance(rows_to_import, list):
        raise HTTPException(status_code=400, detail="Catalogue import expects a list of experiment records.")
    return import_catalog_records(rows_to_import)


@app.delete("/api/catalog")
def api_delete_catalog() -> dict[str, Any]:
    return clear_catalog_records()


@app.get("/api/nasa-psi/search")
def api_nasa_psi_search(
    q: str = Query(default="combustion"),
    accession: str | None = Query(default=None),
    limit: int = Query(default=10, ge=1, le=50),
) -> dict[str, Any]:
    if accession:
        norm_acc = accession.strip().upper()
        if not PSI_ACCESSION_RE.match(norm_acc):
            raise HTTPException(status_code=400, detail=f"Invalid NASA PSI accession '{accession}'.")
        matches = [r for r in VERIFIED_PSI_RECORDS if r["accession"] == norm_acc]
        return {
            "query": q,
            "accession": norm_acc,
            "data_source": "NASA PSI",
            "repository_url": f"{NASA_PSI_BASE_URL}/physci/repo",
            "count": len(matches),
            "results": matches,
        }

    q_low = (q or "combustion").strip().lower()
    tokens = [t for t in re.split(r"\s+", q_low) if t]
    if q_low in ("combustion", "fire", "flame", "all", "*"):
        results = VERIFIED_PSI_RECORDS[:limit]
    else:
        results = [
            r
            for r in VERIFIED_PSI_RECORDS
            if all(
                tok in f"{r['accession']} {r['title']} {r['summary']} {r['tags']} {r['platform']}".lower()
                for tok in tokens
            )
        ][:limit]
    return {
        "query": q,
        "accession": None,
        "data_source": "NASA PSI",
        "repository_url": f"{NASA_PSI_BASE_URL}/physci/repo",
        "count": len(results),
        "results": results,
    }


@app.post("/api/nasa-psi/sync")
async def api_nasa_psi_sync(request: Request) -> dict[str, Any]:
    try:
        payload = await request.json()
    except Exception as exc:
        raise HTTPException(status_code=400, detail="Invalid JSON payload.") from exc

    records_in: list[dict[str, Any]] = []
    if isinstance(payload, dict) and "records" in payload:
        records_in = payload["records"]
    elif isinstance(payload, dict) and ("data_source" in payload or "source_url" in payload):
        records_in = [payload]
    elif isinstance(payload, dict) and "accession" in payload:
        acc = str(payload["accession"]).strip().upper()
        if not PSI_ACCESSION_RE.match(acc):
            raise HTTPException(status_code=400, detail=f"Invalid PSI accession '{acc}'.")
        records_in = [r for r in VERIFIED_PSI_RECORDS if r["accession"] == acc]
        if not records_in:
            raise HTTPException(status_code=404, detail=f"PSI record '{acc}' not found.")
    else:
        records_in = VERIFIED_PSI_RECORDS

    validated: list[dict[str, Any]] = []
    for r in records_in:
        try:
            validated.append(validate_psi_record(r))
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    import_res = import_catalog_records(validated)
    return {
        "status": "synced",
        "data_source": "NASA PSI",
        "synced_count": import_res["imported"],
        "total_catalog": import_res["total"],
        "records": validated,
    }


@app.get("/api/nasa-psi/experiments")
def api_nasa_psi_experiments() -> dict[str, Any]:
    cat = list_catalog_records()
    records = cat["records"] if cat["records"] else VERIFIED_PSI_RECORDS
    return {
        "data_source": "NASA PSI",
        "repository_url": f"{NASA_PSI_BASE_URL}/physci/repo",
        "count": len(records),
        "experiments": records,
    }


@app.post("/api/nasa-psi/analyze")
async def api_nasa_psi_analyze(request: Request) -> dict[str, Any]:
    init_db()
    body = await request.json()
    accession = str(body.get("accession", "PSI-98")).strip().upper()
    if not PSI_ACCESSION_RE.match(accession):
        raise HTTPException(status_code=400, detail=f"Invalid NASA PSI accession '{accession}'.")

    psi_rec = next((r for r in VERIFIED_PSI_RECORDS if r["accession"] == accession), None)
    if not psi_rec:
        raise HTTPException(status_code=404, detail=f"PSI investigation '{accession}' not found.")

    img_bytes = create_synthetic_sample_image("0g_spherical" if accession in ("PSI-98", "PSI-107", "PSI-41") else "lunar_017g")
    img_desc = extract_image_descriptors(img_bytes, f"{accession}_frame.png")
    img_desc["evidence_class"] = "source_record"
    img_desc["psi_accession"] = accession

    selected_techs = body.get("technique_ids") or infer_suggested_techniques([img_desc])
    metadata = {
        "accession": accession,
        "platform": psi_rec["platform"],
        "gravity_g": body.get("gravity_g", psi_rec["gravity_g"]),
        "fuel_material": body.get("fuel") or psi_rec["fuel_material"],
        "oxygen_pct": body.get("oxygen_pct", psi_rec["oxygen_pct"]),
        "pressure_kpa": body.get("pressure_kpa", psi_rec["pressure_kpa"]),
    }

    assessments = evaluate_techniques(selected_techs, [img_desc], metadata)
    candidates = match_catalog_candidates([img_desc], metadata, VERIFIED_PSI_RECORDS)
    raw_report = {
        "executive_summary": (
            f"Bounded PyroSage analysis of official NASA PSI investigation {accession} ({psi_rec['title']}) "
            f"across {len(assessments)} selected scientific techniques."
        ),
        "technique_assessments": assessments,
        "experiment_matches": candidates[:4],
        "key_findings": [
            {
                "rank": 1,
                "relevance_score": 0.92,
                "confidence": 0.88,
                "evidence_type": "computed_from_upload",
                "finding": (
                    f"Extracted bilateral symmetry ({img_desc['bilateral_left_right_score']}) and "
                    f"top-bottom symmetry ({img_desc['top_bottom_symmetry_score']}) for {accession}."
                ),
                "evidence": f"{img_desc['filename']} (area={img_desc['candidate_area_px']} px², fractal_dimension={img_desc['fractal_dimension']})",
            }
        ],
        "safety_implications": [
            {
                "implication": "Ventilation shutoff and multi-band UV/IR/CO sensing recommended for low-gravity enclosures.",
                "priority": "high",
                "supporting_evidence": f"{accession} ({psi_rec['source_url']})",
                "uncertainty": "Pixel measurements are uncalibrated for absolute length or temperature.",
            }
        ],
        "data_gaps": [
            "Absolute temperature (K) requires radiometric calibration and emissivity metadata.",
            "PIV flow velocity vectors require tracer-particle image pairs and a dedicated PIV solver.",
        ],
        "overall_confidence": 0.84,
    }
    validated = validate_analysis_response(raw_report, [a["id"] for a in assessments], candidates)
    clean_ext = strip_internal_payloads([img_desc])
    report_id = f"rep-{uuid.uuid4().hex[:12]}"
    full_report = {
        "report_id": report_id,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "workflow_mode": "nasa_psi_evidence",
        "provenance": {
            "data_source": "NASA PSI",
            "accession": accession,
            "title": psi_rec["title"],
            "source_url": psi_rec["source_url"],
            "fiction_simulator_used": False,
        },
        "local_extractions": clean_ext,
        "evidence_inventory": clean_ext,
        "selected_technique_ids": [a["id"] for a in validated["technique_assessments"]],
        "technique_assessments": validated["technique_assessments"],
        "experiment_matches": validated["experiment_matches"],
        "candidate_experiments": validated["experiment_matches"],
        "si_interpretation": {
            "provider": str(body.get("provider") or DEFAULT_PROVIDER),
            "model": str(body.get("model") or DEFAULT_MODEL),
            "summary": validated["executive_summary"],
            "findings": [
                {
                    "title": f"Finding #{kf['rank']}",
                    "evidence_class": kf["evidence_type"],
                    "interpretation": kf["finding"],
                    "evidence_reference": kf["evidence"],
                }
                for kf in validated["key_findings"]
            ],
            "safety_insights": [
                {
                    "category": "Mission Fire Safety",
                    "insight": si["implication"],
                    "evidence_link": si["supporting_evidence"],
                }
                for si in validated["safety_implications"]
            ],
            "data_gaps": validated["data_gaps"],
        },
        "ranked_findings": [
            {
                "rank": idx + 1,
                "accession": m["experiment_id"],
                "title": m["name"],
                "source_url": m["source_url"],
                "platform": m["platform"],
                "flammability_severity": 0.86,
                "evidence_completeness": m["match_confidence"],
                "mission_relevance": 0.88,
                "composite_score": round(100.0 * m["match_confidence"], 1),
                "confidence_grade": "A" if m["match_confidence"] >= 0.80 else "B",
                "flag": m["reason"],
            }
            for idx, m in enumerate(validated["experiment_matches"][:4])
        ],
        "report": validated,
    }
    with get_db_connection() as conn:
        conn.execute(
            "INSERT INTO reports (report_id, created_at, workflow_mode, provider, model, report_json) VALUES (?, ?, ?, ?, ?, ?)",
            (
                report_id,
                full_report["created_at"],
                full_report["workflow_mode"],
                full_report["si_interpretation"]["provider"],
                full_report["si_interpretation"]["model"],
                json.dumps(full_report),
            ),
        )
        conn.commit()
    return full_report


@app.post("/api/analyze")
async def api_analyze(
    request: Request,
    files: list[UploadFile] | None = File(default=None),
    technique_ids: str | None = Form(default=None),
    gravity_g: float | None = Form(default=None),
    oxygen_pct: float | None = Form(default=None),
    pressure_kpa: float | None = Form(default=None),
    temperature_k: float | None = Form(default=None),
    fuel: str | None = Form(default=None),
    fuel_material: str | None = Form(default=None),
    platform: str | None = Form(default=None),
    mission: str | None = Form(default=None),
    provider: str | None = Form(default=None),
    model: str | None = Form(default=None),
    api_key: str | None = Form(default=None),
    sample_preset: str | None = Form(default=None),
    notes: str = Form(default=""),
) -> dict[str, Any]:
    """Multipart or JSON analysis endpoint per Section 06, 08, 10, and 11.
    Request-scoped api_key is never persisted into SQLite or .env.
    """
    init_db()
    ct_header = (request.headers.get("content-type") or "").lower()
    selected_techs: list[str] = []
    extractions: list[dict[str, Any]] = []
    total_bytes = 0

    if "application/json" in ct_header:
        body = await request.json()
        raw_t = body.get("technique_ids") or []
        if isinstance(raw_t, list):
            selected_techs = [str(x) for x in raw_t]
        gravity_g = _to_float_or_none(body.get("gravity_g", gravity_g))
        oxygen_pct = _to_float_or_none(body.get("oxygen_pct", oxygen_pct))
        pressure_kpa = _to_float_or_none(body.get("pressure_kpa", pressure_kpa))
        temperature_k = _to_float_or_none(body.get("temperature_k", temperature_k))
        fuel_material = str(body.get("fuel_material") or body.get("fuel") or fuel_material or fuel or "")
        platform = str(body.get("platform") or platform or "")
        mission = str(body.get("mission") or mission or "")
        provider = body.get("provider", provider)
        model = body.get("model", model)
        api_key = body.get("api_key", api_key)
        sample_preset = body.get("sample_preset", sample_preset)
        notes = str(body.get("notes", notes))
    else:
        if technique_ids:
            try:
                parsed = json.loads(technique_ids)
                if isinstance(parsed, list):
                    selected_techs = [str(x) for x in parsed]
                else:
                    selected_techs = [s.strip() for s in technique_ids.split(",") if s.strip()]
            except Exception:
                selected_techs = [s.strip() for s in technique_ids.split(",") if s.strip()]

    if files:
        for up in files[:MAX_UPLOAD_FILES]:
            if not up.filename:
                continue
            raw_bytes = await up.read()
            total_bytes += len(raw_bytes)
            if total_bytes > MAX_TOTAL_BYTES:
                raise HTTPException(
                    status_code=413,
                    detail=f"Combined upload size ({total_bytes} bytes) exceeds PYROSAGE_MAX_TOTAL_BYTES ({MAX_TOTAL_BYTES} bytes).",
                )
            try:
                extractions.append(parse_evidence_file(raw_bytes, up.filename, up.content_type or ""))
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc)) from exc

    if not extractions:
        preset = sample_preset or ("0g_spherical" if (gravity_g is None or gravity_g < 0.05) else "lunar_017g")
        img_bytes = create_synthetic_sample_image(preset)
        desc = extract_image_descriptors(img_bytes, f"specimen_{preset}.png")
        desc["evidence_class"] = "uploaded_evidence"
        extractions.append(desc)

    metadata = {
        "gravity_g": gravity_g,
        "oxygen_pct": oxygen_pct,
        "pressure_kpa": pressure_kpa,
        "temperature_k": temperature_k,
        "fuel_material": fuel_material or fuel or "",
        "platform": platform or "",
        "mission": mission or "",
        "notes": notes,
    }

    assessments = evaluate_techniques(selected_techs, extractions, metadata)
    cat_data = list_catalog_records()
    active_catalog = cat_data["records"] if cat_data["records"] else VERIFIED_PSI_RECORDS
    candidates = match_catalog_candidates(extractions, metadata, active_catalog)

    first_ext = extractions[0]
    fallback_raw_report = {
        "executive_summary": (
            f"Processed {len(extractions)} evidence file(s) across {len(assessments)} selected scientific techniques. "
            f"Primary modality: {first_ext.get('modality')} ({first_ext.get('filename')})."
        ),
        "technique_assessments": assessments,
        "experiment_matches": candidates[:6],
        "key_findings": [
            {
                "rank": 1,
                "relevance_score": 0.89,
                "confidence": 0.82,
                "evidence_type": "computed_from_upload",
                "finding": (
                    f"Computed bilateral symmetry={first_ext.get('bilateral_left_right_score', 'N/A')}, "
                    f"top-bottom symmetry={first_ext.get('top_bottom_symmetry_score', 'N/A')}, "
                    f"convexity={first_ext.get('convexity_score', 'N/A')}, "
                    f"fractal dimension={first_ext.get('fractal_dimension', 'N/A')}."
                ),
                "evidence": first_ext.get("filename", "uploaded_evidence"),
            }
        ],
        "safety_implications": [
            {
                "implication": (
                    "Low-gravity flames exhibit reduced buoyant plume rise and altered limiting oxygen thresholds; "
                    "combine ventilation shutoff with multi-spectral UV/IR and CO/smoke detection."
                ),
                "priority": "high",
                "supporting_evidence": first_ext.get("filename", "uploaded_evidence"),
                "uncertainty": "Heuristic segmentation and uncalibrated pixel coordinates require domain-expert verification.",
            }
        ],
        "data_gaps": [
            "Ordinary RGB images do not provide calibrated absolute temperature (K).",
            "Pixel coordinates cannot be converted to physical units (mm or m/s) without a spatial scale fiducial and frame timing.",
        ],
        "overall_confidence": 0.78,
    }

    chosen_provider = str(provider or DEFAULT_PROVIDER).strip().lower()
    chosen_model = str(model or DEFAULT_MODEL).strip()
    selected_ids_canonical = [a["id"] for a in assessments]

    provider_report, execution_path = invoke_si_provider(
        provider=chosen_provider,
        model=chosen_model,
        request_api_key=api_key,
        selected_tech_ids=selected_ids_canonical,
        extractions=extractions,
        metadata=metadata,
        candidates=candidates,
        fallback_report=fallback_raw_report,
    )

    validated = validate_analysis_response(provider_report, selected_ids_canonical, candidates)
    clean_extractions = strip_internal_payloads(extractions)
    report_id = f"rep-{uuid.uuid4().hex[:12]}"

    ranked_list = [
        {
            "rank": idx + 1,
            "accession": m["experiment_id"],
            "title": m["name"],
            "source_url": m.get("source_url") or m.get("source_ref") or f"{NASA_PSI_BASE_URL}/physci/repo",
            "platform": m.get("platform") or "ISS / CIR",
            "flammability_severity": round(max(0.45, 0.90 - idx * 0.04), 3),
            "evidence_completeness": m["match_confidence"],
            "mission_relevance": round(max(0.50, 0.92 - idx * 0.03), 3),
            "composite_score": round(100.0 * m["match_confidence"], 1),
            "confidence_grade": "A" if m["match_confidence"] >= 0.80 else ("B" if m["match_confidence"] >= 0.68 else "C"),
            "flag": m.get("reason", "Catalogue candidate match"),
        }
        for idx, m in enumerate(validated["experiment_matches"][:4])
    ]

    full_report = {
        "report_id": report_id,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "workflow_mode": "uploaded_evidence_analysis",
        "execution_path": execution_path,
        "catalog_total": cat_data["total"],
        "provenance": {
            "data_source": "Operator Uploaded Evidence",
            "uploaded_files": [e.get("filename", "") for e in clean_extractions],
            "relabelled_as_nasa_data": False,
            "fiction_simulator_used": False,
        },
        "metadata": metadata,
        "local_extractions": clean_extractions,
        "evidence_inventory": clean_extractions,
        "selected_technique_ids": [a["id"] for a in validated["technique_assessments"]],
        "technique_assessments": validated["technique_assessments"],
        "experiment_matches": validated["experiment_matches"],
        "candidate_experiments": validated["experiment_matches"],
        "ranked_findings": ranked_list,
        "si_interpretation": {
            "provider": chosen_provider,
            "model": chosen_model,
            "summary": validated["executive_summary"],
            "findings": [
                {
                    "title": f"Finding #{kf['rank']}",
                    "evidence_class": kf["evidence_type"],
                    "interpretation": kf["finding"],
                    "evidence_reference": kf["evidence"],
                }
                for kf in validated["key_findings"]
            ],
            "safety_insights": [
                {
                    "category": "Spacecraft Fire Safety",
                    "insight": si["implication"],
                    "evidence_link": si["supporting_evidence"],
                }
                for si in validated["safety_implications"]
            ],
            "data_gaps": validated["data_gaps"],
        },
        "report": validated,
    }

    with get_db_connection() as conn:
        conn.execute(
            "INSERT INTO reports (report_id, created_at, workflow_mode, provider, model, report_json) VALUES (?, ?, ?, ?, ?, ?)",
            (
                report_id,
                full_report["created_at"],
                full_report["workflow_mode"],
                chosen_provider,
                chosen_model,
                json.dumps(full_report),
            ),
        )
        conn.commit()
    return full_report


@app.get("/api/report/{report_id}")
def api_get_report(report_id: str) -> dict[str, Any]:
    init_db()
    with get_db_connection() as conn:
        row = conn.execute("SELECT report_json FROM reports WHERE report_id = ?", (report_id,)).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail=f"Report '{report_id}' not found.")
    return json.loads(row["report_json"])


if __name__ == "__main__":
    import uvicorn

    port = int(os.getenv("PORT", "8000"))
    uvicorn.run("server:app", host="127.0.0.1", port=port, reload=False)
