"""Automated verification suite for the PyroSage microgravity combustion evidence pipeline."""

from __future__ import annotations

import io
import json
import os
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

try:
    import cv2
except Exception:
    cv2 = None

import numpy as np
import openpyxl

try:
    import pymupdf as fitz
except Exception:
    fitz = None

import pytest
from docx import Document as DocxDocument
from fastapi.testclient import TestClient

import server


@pytest.fixture(autouse=True)
def isolated_db(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    db_file = tmp_path / "test_pyrosage.sqlite3"
    monkeypatch.setenv("PYROSAGE_DB_PATH", str(db_file))
    server.init_db()
    return db_file


@pytest.fixture()
def client() -> TestClient:
    return TestClient(server.app)


def test_1_technique_registry_has_57_methods_and_canonical_ids() -> None:
    """Section 12 Test 1: Confirms 57 untouched core methods + 20 extended differentiation methods (77 total), 9 groups, and key canonical IDs."""
    techniques = server.load_techniques()
    assert len(techniques) == 77

    ids = [t["id"] for t in techniques]
    assert len(set(ids)) == 77
    for key_id in (
        "absence_buoyancy",
        "spherical_symmetry",
        "piv",
        "segmentation",
        "rotational_symmetry",
        "fractal_dimension",
        "convexity_defects",
        "cnn_resnet",
        "symmetry_loss",
        "spectral_deconvolution",
        "chemical_kinetics",
        "gravity_scaling",
        "lunar_gravity",
        "martian_gravity",
        "cross_experiment_compare",
        "evidence_ranking",
        "anomaly_detection",
        "risk_interpretation",
        "terrestrial_vs_microgravity",
        "multi_gravity_behavior",
        "spherical_vs_teardrop",
        "flame_regime_classification",
        "lunar_martian_predictive_safety",
    ):
        assert key_id in ids

    families = {t["group"] for t in techniques}
    expected_families = {
        "Microgravity combustion fundamentals",
        "Optical diagnostics & flow",
        "Thermal & cool-flame diagnostics",
        "Platforms & cross-platform comparison",
        "Geometry & object classification",
        "Numerical & combustion measurements",
        "Chemical kinetics & spectroscopy",
        "Gravity scaling & partial gravity",
        "SI interpretation & safety assessment",
    }
    assert families == expected_families

    for tech in techniques:
        assert tech.get("name")
        assert tech.get("description")
        assert isinstance(tech.get("modalities"), list) and len(tech["modalities"]) >= 1
        assert tech.get("prerequisites")


def test_2_catalogue_round_trip_empty_import_list_and_clear(client: TestClient) -> None:
    """Section 12 Test 2: Confirms catalogue starts empty, imports records, lists them, and clears back to 0."""
    initial = client.get("/api/catalog")
    assert initial.status_code == 200
    assert initial.json()["total"] == 0
    assert initial.json()["records"] == []

    sample_records = [
        {
            "experiment_id": "CAT-FLEX-01",
            "name": "FLEX Droplet Test 01",
            "mission": "ISS CIR",
            "platform": "ISS / CIR",
            "gravity_g": 0.0,
            "fuel_material": "heptane",
            "oxygen_pct": 21.0,
            "pressure_kpa": 101.3,
            "temperature_k": 298.0,
            "techniques": "segmentation, spherical_symmetry",
            "summary": "0g spherical heptane burn",
            "bilateral_left_right_score": 0.96,
            "top_bottom_symmetry_score": 0.95,
            "convexity_score": 0.98,
            "fractal_dimension": 1.07,
        },
        {
            "title": "BASS-II Solid PMMA Rod",
            "facility": "ISS MSG",
            "gravity": 0.0,
            "fuel": "pmma",
            "oxygen": 18.5,
        },
    ]
    imp = client.post("/api/catalog/import", json={"experiments": sample_records})
    assert imp.status_code == 200
    imp_data = imp.json()
    assert imp_data["imported"] == 2
    assert imp_data["total"] == 2

    listed = client.get("/api/catalog")
    assert listed.status_code == 200
    assert listed.json()["total"] == 2
    ids = [r["experiment_id"] for r in listed.json()["records"]]
    assert "CAT-FLEX-01" in ids
    assert any(i.startswith("CAT-") for i in ids)

    cleared = client.delete("/api/catalog")
    assert cleared.status_code == 200
    assert cleared.json()["deleted"] == 2
    assert cleared.json()["total"] == 0

    after_clear = client.get("/api/catalog")
    assert after_clear.status_code == 200
    assert after_clear.json()["total"] == 0


def test_3_image_extraction_computes_symmetry_and_rotation_features() -> None:
    """Section 12 Test 3: Generates synthetic flame image and confirms symmetry and rotation features."""
    spherical_bytes = server.create_synthetic_sample_image("0g_spherical")
    plume_bytes = server.create_synthetic_sample_image("martian_038g")

    sph_desc = server.extract_image_descriptors(spherical_bytes, "spherical_0g.png")
    plume_desc = server.extract_image_descriptors(plume_bytes, "plume_038g.png")

    assert sph_desc["modality"] == "image"
    assert sph_desc["candidate_area_px"] > 100.0
    assert sph_desc["candidate_perimeter_px"] > 20.0
    assert sph_desc["bilateral_left_right_score"] > 0.85
    assert sph_desc["top_bottom_symmetry_score"] > plume_desc["top_bottom_symmetry_score"]
    assert sph_desc["reflection_iou"] > 0.80
    assert "rotational_symmetry_iou" in sph_desc
    assert set(sph_desc["rotational_symmetry_iou"].keys()) == {"deg_90", "deg_120", "deg_180", "mean"}
    assert sph_desc["rotational_symmetry"] > 0.75
    assert 0.80 <= sph_desc["convexity_score"] <= 1.0
    assert 1.0 <= sph_desc["fractal_dimension"] < 2.0
    assert sph_desc["peak_count"] >= 1
    assert len(sph_desc["geometric_centroid"]) == 2
    assert len(sph_desc["intensity_centroid"]) == 2


def test_4_response_validation_removes_unselected_and_invented_ids_and_locks_piv() -> None:
    """Section 12 Test 4: Confirms validate_analysis_response() removes unselected techniques, invented experiment IDs, and forces piv to requires_more_data."""
    candidates = [
        {
            "experiment_id": "PSI-98",
            "accession": "PSI-98",
            "name": "FLEX",
            "pre_match_score": 0.91,
        }
    ]
    raw_report = {
        "executive_summary": "Test report",
        "technique_assessments": [
            {"technique_id": "segmentation", "status": "computed_from_upload", "observations": "Mask computed."},
            {"technique_id": "piv", "status": "computed_from_upload", "observations": "Fake PIV claim."},
            {"technique_id": "cnn_resnet", "status": "visually_assessed", "observations": "Fake CNN claim."},
            {"technique_id": "raman", "status": "computed_from_upload", "observations": "Unselected method."},
        ],
        "experiment_matches": [
            {"experiment_id": "PSI-98", "match_confidence": 1.4, "reason": "Valid candidate"},
            {"experiment_id": "PSI-99999", "match_confidence": 0.95, "reason": "Invented experiment ID"},
        ],
        "key_findings": [
            {
                "rank": 1,
                "relevance_score": 1.5,
                "confidence": -0.2,
                "evidence_type": "computed_from_upload",
                "technique_ids": ["segmentation", "piv", "raman"],
                "finding": "Compared PSI-98 and hallucinated PSI-88888.",
                "evidence": "spherical_0g.png",
            }
        ],
        "data_gaps": [],
        "overall_confidence": 1.25,
    }

    validated = server.validate_analysis_response(
        raw_report=raw_report,
        selected_technique_ids=["segmentation", "piv", "cnn_resnet", "bilateral_symmetry"],
        candidate_experiments=candidates,
    )

    tech_by_id = {t["technique_id"]: t for t in validated["technique_assessments"]}
    assert set(tech_by_id.keys()) == {"segmentation", "piv", "cnn_resnet", "bilateral_symmetry"}
    assert "raman" not in tech_by_id
    # piv and cnn_resnet must be forced to requires_more_data
    assert tech_by_id["piv"]["status"] == "requires_more_data"
    assert tech_by_id["cnn_resnet"]["status"] == "requires_more_data"
    # missing selected technique bilateral_symmetry supplied with requires_more_data
    assert tech_by_id["bilateral_symmetry"]["status"] == "requires_more_data"

    # Invented experiment ID removed; valid candidate retained with clamped confidence
    matched_ids = [m["experiment_id"] for m in validated["experiment_matches"]]
    assert matched_ids == ["PSI-98"]
    assert validated["experiment_matches"][0]["match_confidence"] == 1.0
    assert "PSI-99999" in validated["rejected_experiment_ids"]
    assert "PSI-88888" in validated["rejected_experiment_ids"]
    assert "PSI-88888" not in validated["key_findings"][0]["finding"]
    assert validated["key_findings"][0]["relevance_score"] == 1.0
    assert validated["key_findings"][0]["confidence"] == 0.0
    assert validated["overall_confidence"] == 1.0


def test_5_geometry_matching_scores_matching_catalogue_record_higher() -> None:
    """Section 12 Test 5: Confirms imported optional geometry descriptors match an uploaded geometry descriptor."""
    spherical_bytes = server.create_synthetic_sample_image("0g_spherical")
    img_desc = server.extract_image_descriptors(spherical_bytes, "droplet_0g.png")

    catalog_rows = [
        {
            "experiment_id": "CAT-SPHERICAL",
            "name": "Spherical 0g Droplet Reference",
            "platform": "ISS / CIR",
            "gravity_g": 0.0,
            "fuel_material": "heptane",
            "bilateral_left_right_score": img_desc["bilateral_left_right_score"],
            "top_bottom_symmetry_score": img_desc["top_bottom_symmetry_score"],
            "convexity_score": img_desc["convexity_score"],
            "fractal_dimension": img_desc["fractal_dimension"],
        },
        {
            "experiment_id": "CAT-ASYMMETRIC",
            "name": "Asymmetric 1g Upward Plume Reference",
            "platform": "1g Rig",
            "gravity_g": 1.0,
            "fuel_material": "cotton",
            "bilateral_left_right_score": 0.35,
            "top_bottom_symmetry_score": 0.20,
            "convexity_score": 0.45,
            "fractal_dimension": 1.75,
        },
    ]
    server.import_catalog_records(catalog_rows)
    matches = server.match_catalog_candidates([img_desc], {"gravity_g": 0.0, "fuel_material": "heptane"})

    assert len(matches) == 2
    assert matches[0]["experiment_id"] == "CAT-SPHERICAL"
    assert matches[0]["pre_match_score"] > matches[1]["pre_match_score"]
    assert "geometry descriptor similarity" in matches[0]["reason"]


def test_video_table_and_document_extractors(tmp_path: Path) -> None:
    # Video
    video_path = tmp_path / "flame_test.avi"
    fourcc = cv2.VideoWriter_fourcc(*"MJPG")
    writer = cv2.VideoWriter(str(video_path), fourcc, 15.0, (160, 160))
    for idx in range(12):
        frame = np.zeros((160, 160, 3), dtype=np.uint8)
        radius = 30 + int(4 * np.sin(idx * 0.8))
        cv2.circle(frame, (80 + (idx % 3), 80), radius, (30, 120, 240), -1)
        writer.write(frame)
    writer.release()

    v_desc = server.extract_video_descriptors(video_path.read_bytes(), "flame_test.avi")
    assert v_desc["modality"] == "video"
    assert v_desc["sampled_frames"] >= 3
    assert v_desc["piv_computed"] is False

    # Spectral CSV
    csv_spectrum = (
        "wavelength_nm,intensity_counts\n"
        "300.0,12.0\n"
        "308.5,185.0\n"
        "320.0,15.0\n"
    ).encode("utf-8")
    t_desc = server.extract_table_descriptors(csv_spectrum, "oh_spectrum.csv")
    assert t_desc["modality"] == "spectral"
    assert t_desc["spectral_analysis"]["dominant_wavelength_or_band"] == pytest.approx(308.5)

    # PDF, DOCX, XLSX
    pdf_doc = fitz.open()
    page = pdf_doc.new_page()
    page.insert_text((72, 72), "NASA PSI-98 FLEX heptane microgravity combustion.")
    pdf_bytes = pdf_doc.tobytes()
    pdf_doc.close()
    pdf_desc = server.extract_document_descriptors(pdf_bytes, "flex.pdf")
    assert "PSI-98" in pdf_desc["detected_psi_accessions"]

    docx_obj = DocxDocument()
    docx_obj.add_paragraph("SoFIE PSI-112 PMMA test.")
    docx_buf = io.BytesIO()
    docx_obj.save(docx_buf)
    docx_desc = server.extract_document_descriptors(docx_buf.getvalue(), "sofie.docx")
    assert "PSI-112" in docx_desc["detected_psi_accessions"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Accession", "Fuel"])
    ws.append(["PSI-25", "PMMA"])
    xlsx_buf = io.BytesIO()
    wb.save(xlsx_buf)
    xlsx_desc = server.extract_document_descriptors(xlsx_buf.getvalue(), "bass2.xlsx")
    assert "PSI-25" in xlsx_desc["detected_psi_accessions"]


def test_end_to_end_api_routes_and_nasa_psi_connector(client: TestClient) -> None:
    health = client.get("/api/health")
    assert health.status_code == 200
    assert health.json()["technique_count"] == 77
    assert health.json()["catalog_count"] == 0

    techs = client.get("/api/techniques", params={"group": "Geometry & object classification"})
    assert techs.status_code == 200
    assert techs.json()["count"] == 14

    sample_spec = client.get("/api/sample-specimen/0g_spherical")
    assert sample_spec.status_code == 200
    assert sample_spec.json()["preset"] == "0g_spherical"
    assert sample_spec.json()["descriptors"]["bilateral_left_right_score"] > 0.85

    psi_search = client.get("/api/nasa-psi/search", params={"accession": "PSI-98"})
    assert psi_search.status_code == 200
    assert psi_search.json()["results"][0]["accession"] == "PSI-98"

    bad_psi = client.get("/api/nasa-psi/search", params={"accession": "INVALID-123"})
    assert bad_psi.status_code == 400

    img_bytes = server.create_synthetic_sample_image("0g_spherical")
    upload_res = client.post(
        "/api/analyze",
        files=[("files", ("droplet_0g.png", img_bytes, "image/png"))],
        data={
            "technique_ids": json.dumps(["segmentation", "spherical_symmetry", "bilateral_symmetry", "piv", "risk_interpretation", "terrestrial_vs_microgravity"]),
            "gravity_g": "0.0",
            "fuel": "heptane",
        },
    )
    assert upload_res.status_code == 200
    rep = upload_res.json()
    statuses = {a["technique_id"]: a["status"] for a in rep["technique_assessments"]}
    assert statuses["segmentation"] == "computed_from_upload"
    assert statuses["terrestrial_vs_microgravity"] == "computed_from_upload"
    assert statuses["piv"] == "requires_more_data"
    assert "nearest_psi_twin" in rep and rep["nearest_psi_twin"]["accession"].startswith("PSI-")
    assert "moon_mars_scaling" in rep and len(rep["moon_mars_scaling"]["regimes"]) == 4
    seg_entry = next(a for a in rep["technique_assessments"] if a["technique_id"] == "segmentation")
    assert "differentiation" in seg_entry and seg_entry["differentiation"]
    assert isinstance(seg_entry.get("numeric_score"), (int, float))

    fetched = client.get(f"/api/report/{rep['report_id']}")
    assert fetched.status_code == 200
    assert fetched.json()["report_id"] == rep["report_id"]

    dash = client.get("/")
    assert dash.status_code == 200
    assert "PyroSage" in dash.text

    main_dash = client.get("/pyrosage_interactive_dashboard.html")
    assert main_dash.status_code == 200
    assert "PyroSage" in main_dash.text
    assert server.MAIN_DASHBOARD_PATH.exists()

    proto = client.get("/reference_prototype.html")
    assert proto.status_code == 200
    assert "Fiction" in proto.text

    connector_doc = server.ROOT_DIR / "docs" / "NASA_PSI_CONNECTOR.md"
    tech_pdf = server.ROOT_DIR / "docs" / "PyroSage_Technical_Documentation.pdf"
    assert connector_doc.exists() and connector_doc.stat().st_size > 500
    assert tech_pdf.exists() and tech_pdf.stat().st_size > 100000
    with fitz.open(str(tech_pdf)) as pdf_doc:
        assert len(pdf_doc) == 21


def test_provider_schema_and_image_payload_limits() -> None:
    """Section 06 & 08: Verify PDF 3-page rendering, 5-image payload cap, and 9 mandatory SI system prompt rules."""
    pdf_doc = fitz.open()
    for p_num in range(4):
        page = pdf_doc.new_page()
        page.insert_text((72, 72), f"Page {p_num + 1} PSI-98 FLEX droplet data.")
    pdf_bytes = pdf_doc.tobytes()
    pdf_doc.close()

    pdf_desc = server.extract_document_descriptors(pdf_bytes, "four_pages.pdf")
    assert pdf_desc["page_count"] == 4
    assert pdf_desc["rendered_page_count"] == 3

    img_desc = server.extract_image_descriptors(server.create_synthetic_sample_image("0g_spherical"), "flame.png")
    payloads = server.collect_image_payloads([pdf_desc, img_desc, img_desc, img_desc])
    assert len(payloads) == server.MAX_IMAGE_PAYLOADS == 5

    tech_map = server.get_technique_map()
    sys_prompt = server.build_si_system_instruction(
        [tech_map["segmentation"], tech_map["piv"]],
        [{"experiment_id": "PSI-98", "name": "FLEX", "platform": "ISS / CIR", "gravity_g": 0.0, "pre_match_score": 0.92}],
    )
    assert "MANDATORY ANALYSIS RULES" in sys_prompt
    assert "Optical flow is not PIV" in sys_prompt
    assert "segmentation" in sys_prompt
    assert "PSI-98" in sys_prompt
    assert set(server.SI_RESPONSE_JSON_SCHEMA["required"]) == {
        "executive_summary",
        "technique_assessments",
        "experiment_matches",
        "key_findings",
        "safety_implications",
        "data_gaps",
        "overall_confidence",
    }

