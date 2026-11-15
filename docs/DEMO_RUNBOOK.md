# PyroSage — Hackathon Demo Runbook

## Demo objective
Demonstrate an evidence-first workflow for discovering and interpreting microgravity combustion research. Keep the fictional simulator visibly separate from the NASA evidence workflow.

## Suggested 5–7 minute walkthrough
1. **Problem framing (30 seconds)** — Explain why historic microgravity flame experiments are difficult to discover, compare, and translate into crew-safety insight.
2. **Dashboard orientation (30 seconds)** — Show the separate Fiction Combustion and Evidence Analysis workspaces.
3. **NASA PSI discovery (60 seconds)** — Search the official NASA Physical Sciences Informatics repository or open a known PSI investigation. Point out accession, investigation metadata, source links, and provenance.
4. **Evidence intake (60 seconds)** — Upload a representative image, video, table, or experiment document. Explain supported file types and file-size settings.
5. **Technique selection (60 seconds)** — Select only methods appropriate to the uploaded modality. Highlight prerequisites for PIV, spectroscopy, thermal imaging, and trained classifiers.
6. **SI analysis (90 seconds)** — Run the configured Gemini or OpenAI provider. Show the selected technique list, extracted descriptors, source experiment candidates, and structured result.
7. **Evidence review (60 seconds)** — Explain which outputs are measured, inferred, unavailable, or require specialist instrumentation. Open the report and export JSON.
8. **Close (30 seconds)** — Connect evidence-backed findings to fire-safety questions for lunar and Mars missions.

## Before presenting
- Install dependencies and run the project locally.
- Configure a provider key in the local .env file; never paste keys into slides, source control, or a public screen.
- Check the selected provider/model setting and test one small request.
- Confirm the NASA PSI repository is reachable from the demo network.
- Prepare one small, rights-cleared sample asset with a known source investigation.
- Keep the experiment catalogue populated only with records retrieved from NASA PSI.
- Run the automated test suite and confirm it passes.
- Close unrelated tabs and hide local environment files.

## Demo integrity
- Do not present the Fiction Combustion simulator as NASA experimental evidence.
- Do not describe a technique as executed when its required data or solver is unavailable.
- Do not present model-generated interpretations as NASA conclusions.
- Show the original PSI source and accession alongside every matched investigation.
- Explain uncertainty and missing measurements in plain language.

## Recovery plan
- If live PSI search is slow, use a known PSI accession lookup and show its official investigation page.
- If the provider is unavailable, demonstrate local feature extraction and the report structure without claiming a live model result.
- If an uploaded file is unsupported, switch to a CSV or image sample prepared before the demo.
