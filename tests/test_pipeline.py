"""Automated verification suite for the PyroSage microgravity combustion evidence pipeline."""

from __future__ import annotations

import io
import json
from pathlib import Path

import cv2
import numpy as np
import openpyxl
import pymupdf as fitz
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


def test_techniques_registry_has_57_methods_and_9_families() -> None:
    techniques = server.load_techniques()
    assert len(techniques) == 57

    ids = [t["id"] for t in techniques]
    assert len(set(ids)) == 57
    assert ids[0] == "T01"
    assert ids[-1] == "T57"

    families = {t["group"] for t in techniques}
    expected_families = {
        "Microgravity fundamentals",
        "Geometry and vision",
        "Temporal dynamics",
        "Diagnostics",
        "Spectroscopy and chemistry",
        "Platforms",
        "Gravity scaling",
        "AI methods",
        "Safety interpretation",
    }
    assert families == expected_families

    for tech in techniques:
        assert tech.get("name")
        assert tech.get("description")
        assert isinstance(tech.get("supported_modalities"), list) and len(tech["supported_modalities"]) >= 1
        assert isinstance(tech.get("prerequisites"), list) and len(tech["prerequisites"]) >= 1


def test_image_descriptor_extraction_distinguishes_spherical_and_elongated_flames() -> None:
    spherical_bytes = server.create_synthetic_sample_image("0g_spherical")
    plume_bytes = server.create_synthetic_sample_image("martian_038g")

    sph_desc = server.extract_image_descriptors(spherical_bytes, "spherical_0g.png")
    plume_desc = server.extract_image_descriptors(plume_bytes, "plume_038g.png")

    assert sph_desc["modality"] == "image"
    assert sph_desc["spherical_symmetry_score"] > plume_desc["spherical_symmetry_score"]
    assert sph_desc["top_bottom_symmetry"] > plume_desc["top_bottom_symmetry"]
    assert 0.85 <= sph_desc["aspect_ratio"] <= 1.15
    assert 1.0 <= sph_desc["fractal_dimension"] < 2.0
    assert sph_desc["solidity"] > 0.80
    assert sph_desc["convexity"] > 0.80
    assert "units" in sph_desc


def test_video_descriptor_extraction_on_synthetic_sequence(tmp_path: Path) -> None:
    video_path = tmp_path / "flame_test.avi"
    fourcc = cv2.VideoWriter_fourcc(*"MJPG")
    writer = cv2.VideoWriter(str(video_path), fourcc, 15.0, (160, 160))
    for idx in range(12):
        frame = np.zeros((160, 160, 3), dtype=np.uint8)
        radius = 30 + int(4 * np.sin(idx * 0.8))
        cx = 80 + (idx % 3)
        cv2.circle(frame, (cx, 80), radius, (60, 130, 240), -1)
        writer.write(frame)
    writer.release()

    video_bytes = video_path.read_bytes()
    v_desc = server.extract_video_descriptors(video_bytes, "flame_test.avi")
    assert v_desc["modality"] == "video"
    assert v_desc["sampled_frames"] >= 6
    assert v_desc["area_cv"] >= 0.0
    assert v_desc["centroid_jitter_px"] >= 0.0
    assert v_desc["dominant_flicker_hz"] >= 0.0
    assert isinstance(v_desc["extinction_detected"], bool)


def test_table_and_spectrum_parsing() -> None:
    csv_spectrum = (
        "wavelength_nm,intensity_counts,temperature_proxy\n"
        "300.0,12.0,1.0\n"
        "308.5,185.0,2.5\n"
        "320.0,15.0,1.1\n"
        "431.0,140.0,2.2\n"
        "516.0,95.0,1.8\n"
    ).encode("utf-8")

    desc = server.extract_table_descriptors(csv_spectrum, "chemiluminescence_spectrum.csv")
    assert desc["modality"] == "spectrum"
    assert desc["row_count"] == 5
    assert "wavelength_nm" in desc["numeric_summaries"]
    spec = desc["spectral_analysis"]
    assert spec["dominant_wavelength_or_band"] == pytest.approx(308.5)
    assert spec["spectral_snr"] > 1.0
    assert any("OH*" in band for band in spec["candidate_species_bands"])


def test_multimodal_document_extractors_pdf_docx_xlsx_json() -> None:
    # PDF
    pdf_doc = fitz.open()
    page = pdf_doc.new_page()
    page.insert_text((72, 72), "NASA PSI-98 FLEX and PSI-25 BASS-II heptane and PMMA microgravity combustion.")
    pdf_bytes = pdf_doc.tobytes()
    pdf_doc.close()

    pdf_desc = server.extract_document_descriptors(pdf_bytes, "report.pdf")
    assert pdf_desc["page_count"] == 1
    assert "PSI-98" in pdf_desc["detected_psi_accessions"]
    assert "PSI-25" in pdf_desc["detected_psi_accessions"]
    assert "heptane" in pdf_desc["detected_fuels"]

    # DOCX
    docx_obj = DocxDocument()
    docx_obj.add_paragraph("SoFIE PSI-112 investigation on cast PMMA in exploration atmosphere.")
    docx_buf = io.BytesIO()
    docx_obj.save(docx_buf)
    docx_desc = server.extract_document_descriptors(docx_buf.getvalue(), "notes.docx")
    assert "PSI-112" in docx_desc["detected_psi_accessions"]
    assert "pmma" in docx_desc["detected_fuels"]

    # XLSX
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BASS2"
    ws.append(["Test", "Accession", "Fuel", "O2_pct"])
    ws.append([1, "PSI-25", "PMMA", 18.5])
    xlsx_buf = io.BytesIO()
    wb.save(xlsx_buf)
    xlsx_desc = server.extract_document_descriptors(xlsx_buf.getvalue(), "matrix.xlsx")
    assert xlsx_desc["modality"] == "table"
    assert "PSI-25" in xlsx_desc["detected_psi_accessions"]

    # JSON
    json_bytes = json.dumps({"investigation": "PSI-64", "fuel": "cotton", "spectrum": [1, 2, 3]}).encode("utf-8")
    json_desc = server.extract_document_descriptors(json_bytes, "metadata.json")
    assert json_desc["modality"] == "spectrum"
    assert "PSI-64" in json_desc["detected_psi_accessions"]


def test_technique_execution_states_prevent_false_specialist_claims() -> None:
    img_bytes = server.create_synthetic_sample_image("0g_spherical")
    img_desc = server.extract_image_descriptors(img_bytes, "frame.png")

    selected = ["T08", "T09", "T21", "T26", "T28", "T30", "T47"]
    assessments = server.evaluate_techniques(selected, [img_desc], {"gravity_g": 0.0})
    by_id = {a["id"]: a for a in assessments}

    assert by_id["T08"]["state"] == "Executed"
    assert by_id["T09"]["state"] == "Executed"
    # T21 (flicker) requires video or time-series table, not a single still image
    assert by_id["T21"]["state"] == "Requires data"
    # T26 (PIV), T28 (calibrated FLIR), T30 (LIF), T47 (pretrained CNN/ResNet) require specialist solvers
    assert by_id["T26"]["state"] == "Requires specialist solver"
    assert by_id["T28"]["state"] == "Requires specialist solver"
    assert by_id["T30"]["state"] == "Requires specialist solver"
    assert by_id["T47"]["state"] == "Requires specialist solver"


def test_nasa_psi_provenance_validation_and_rejection(client: TestClient) -> None:
    # Valid PSI search
    res = client.get("/api/nasa-psi/search", params={"accession": "PSI-98"})
    assert res.status_code == 200
    data = res.json()
    assert data["count"] == 1
    assert data["results"][0]["accession"] == "PSI-98"
    assert data["results"][0]["data_source"] == "NASA PSI"

    # Reject non-PSI accession format
    bad_acc = client.get("/api/nasa-psi/search", params={"accession": "FAKE-123"})
    assert bad_acc.status_code == 400

    # Reject non-NASA PSI source in /api/nasa-psi/sync
    bad_sync = client.post(
        "/api/nasa-psi/sync",
        json={
            "accession": "PSI-999",
            "title": "Third-Party Dataset",
            "data_source": "External Lab",
            "source_url": "https://example.com/PSI-999",
        },
    )
    assert bad_sync.status_code == 400


def test_si_output_validator_strips_hallucinated_psi_ids() -> None:
    assessments = [
        {"id": "T08", "state": "Executed"},
        {"id": "T26", "state": "Requires specialist solver"},
    ]
    raw_llm = {
        "summary": "Analysis summary",
        "referenced_psi_accessions": ["PSI-98", "PSI-99999"],
        "findings": [
            {
                "title": "Finding with unverified ID PSI-88888",
                "evidence_class": "model_interpretation",
                "technique_ids": ["T08", "T26", "T99"],
                "evidence_reference": "Frame 1",
                "interpretation": "Compared against PSI-98 and hallucinated PSI-88888.",
                "claimed_executed": True,
            }
        ],
        "safety_insights": [],
        "hypotheses": [],
        "data_gaps": [],
    }

    validated = server.validate_si_output(
        raw_output=raw_llm,
        allowed_technique_ids=["T08", "T26"],
        allowed_psi_accessions=["PSI-98", "PSI-25"],
        technique_assessments=assessments,
    )
    assert validated["referenced_psi_accessions"] == ["PSI-98"]
    assert "PSI-99999" in validated["rejected_hallucinated_accessions"]
    assert "PSI-88888" in validated["rejected_hallucinated_accessions"]
    assert validated["findings"][0]["technique_ids"] == ["T08", "T26"]
    assert "PSI-88888" not in validated["findings"][0]["interpretation"]


def test_end_to_end_api_routes_and_report_retrieval(client: TestClient) -> None:
    health = client.get("/api/health")
    assert health.status_code == 200
    assert health.json()["technique_count"] == 57

    techs = client.get("/api/techniques", params={"group": "Geometry and vision"})
    assert techs.status_code == 200
    assert techs.json()["count"] == 10

    exps = client.get("/api/nasa-psi/experiments")
    assert exps.status_code == 200
    assert exps.json()["count"] >= 6

    psi_an = client.post(
        "/api/nasa-psi/analyze",
        json={"accession": "PSI-98", "technique_ids": ["T01", "T04", "T08", "T09", "T52", "T53"]},
    )
    assert psi_an.status_code == 200
    psi_rep = psi_an.json()
    assert psi_rep["provenance"]["accession"] == "PSI-98"
    assert psi_rep["provenance"]["fiction_simulator_used"] is False

    # Upload an image to /api/analyze
    img_bytes = server.create_synthetic_sample_image("0g_spherical")
    upload_res = client.post(
        "/api/analyze",
        files=[("files", ("droplet_0g.png", img_bytes, "image/png"))],
        data={
            "technique_ids": json.dumps(["T01", "T04", "T08", "T09", "T10", "T12", "T14", "T26", "T52", "T53"]),
            "gravity_g": "0.0",
            "oxygen_pct": "21.0",
            "pressure_kpa": "101.3",
            "fuel": "heptane",
        },
    )
    assert upload_res.status_code == 200
    rep = upload_res.json()
    assert rep["provenance"]["relabelled_as_nasa_data"] is False
    assert len(rep["ranked_findings"]) >= 1

    # Retrieve stored report
    fetched = client.get(f"/api/report/{rep['report_id']}")
    assert fetched.status_code == 200
    assert fetched.json()["report_id"] == rep["report_id"]

    # Dashboard & prototype routes and required documentation artifacts
    dash = client.get("/")
    assert dash.status_code == 200
    assert "PyroSage" in dash.text

    proto = client.get("/reference_prototype.html")
    assert proto.status_code == 200
    assert "Fiction" in proto.text

    connector_doc = server.ROOT_DIR / "docs" / "NASA_PSI_CONNECTOR.md"
    tech_pdf = server.ROOT_DIR / "docs" / "PyroSage_Technical_Documentation.pdf"
    assert connector_doc.exists() and connector_doc.stat().st_size > 500
    assert tech_pdf.exists() and tech_pdf.stat().st_size > 5000
    with fitz.open(str(tech_pdf)) as pdf_doc:
        assert len(pdf_doc) >= 4
        assert "PyroSage" in pdf_doc[0].get_text("text")

