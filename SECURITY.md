# Security Policy

## Reporting a vulnerability

Please do not publish API keys, private research data, or exploit details in a public issue. Report security concerns privately to the repository maintainer through the contact method configured in the GitHub profile.

## Secret handling

- Store model provider keys in a local `.env` file or deployment secret manager.
- Never commit `.env`, credentials, access tokens, or private datasets.
- Rotate any key that is accidentally exposed.
- Avoid logging authorization headers, provider keys, or sensitive upload contents.

## Data handling

PyroSage keeps local evidence processing and its SQLite cache under the operator's control. Review data-sharing settings before sending evidence to an external model provider.
