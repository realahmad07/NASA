# Contributing to PyroSage

Thank you for helping improve PyroSage. Contributions should preserve scientific traceability, data provenance, and the separation between measured evidence and model-generated interpretation.

## Development setup

1. Fork the repository and create a feature branch.
2. Create and activate a Python virtual environment.
3. Install dependencies with `pip install -r requirements.txt` and `pip install pytest`.
4. Copy `.env.example` to `.env` and add provider keys locally when API testing is needed.
5. Run `python -m pytest -q` before submitting changes.

## Contribution principles

- Keep NASA PSI provenance explicit. Do not mix third-party records into the NASA-only experiment catalogue.
- Add or update tests for parsers, provenance rules, technique selection, and output validation.
- Do not describe a technique as executed unless its required input data and computation are available.
- Preserve units, instrument metadata, calibration context, source URLs, and uncertainty where supplied.
- Never commit API keys, personal data, private datasets, local databases, or generated credentials.
- Keep Fiction Combustion separate from archival-evidence analysis.

## Pull request checklist

- [ ] Tests pass locally.
- [ ] New technique records include modalities and prerequisites.
- [ ] NASA PSI data retain source identifiers and URLs.
- [ ] API responses are validated before rendering or storage.
- [ ] Documentation reflects submitted code.
