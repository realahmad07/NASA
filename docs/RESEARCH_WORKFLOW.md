# NASA PSI Evidence Workflow

## Workflow
~~~mermaid
flowchart TD
  A[NASA PSI search] --> B[Open investigation]
  B --> C[Verify accession and source]
  C --> D[Select linked NASA assets]
  D --> E[Inspect modality and metadata]
  E --> F[Select applicable techniques]
  F --> G[Run local extractors]
  G --> H[Match against PSI catalogue]
  H --> I[Send bounded context to LLM]
  I --> J[Validate structured response]
  J --> K[Review and export report]
~~~

## Evidence-to-technique routing
| Evidence modality | Suitable analysis families | Data needed for stronger conclusions |
|---|---|---|
| Visible image | Segmentation, shape, symmetry, centroid, convexity, boundary geometry | Scale calibration, exposure metadata, reliable flame mask |
| Video | Frame-wise geometry, centroid motion, flicker and temporal variation | Frame rate, timestamps, stable camera, adequate temporal coverage |
| Thermal image | Temperature-field and thermal-gradient analysis | Radiometric calibration, emissivity, camera metadata |
| Particle images/video | PIV and flow-field reconstruction | Tracer particles, image pairs, timing interval, calibration |
| Spectral data | LIF, FTIR, Raman and species analysis | Wavelength axis, calibration, reference spectra, instrument metadata |
| Numerical tables | Statistical comparison, trends, correlations, gravity/platform comparison | Units, uncertainty, sample definitions, experimental conditions |
| Experiment documents | Metadata extraction, evidence linking, literature comparison | Traceable source and page/section references |

## Experiment matching
Matching should combine only fields that are present and comparable:
- PSI accession and investigation metadata
- Platform and mission context
- Gravity level and exposure duration
- Fuel, oxidizer, geometry, and burner configuration
- Available diagnostics and sensor modality
- Computed flame descriptors, where image calibration supports comparison

A similarity score is a retrieval aid, not proof that two experiments are scientifically equivalent.

## Report contents
1. Uploaded evidence inventory.
2. NASA PSI source records and accession IDs.
3. Selected techniques and execution states.
4. Computed descriptors with units and method details.
5. Candidate experiments and match rationale.
6. Model interpretation tied to evidence references.
7. Safety relevance and operational implications.
8. Uncertainty, missing measurements, and follow-up experiments.
