# Automated Engineering Drawing Digitisation Pipeline

A Python automation pipeline that uses a locally-hosted vision-language model (VLM) to extract structured tabular data from AutoCAD-style fabrication drawing screenshots — built to solve a real manual data-entry bottleneck during a live structural steel construction QC workflow.

## Background

During a Practice School-I internship in structural steel fabrication and quality control at Kalpataru Projects International Ltd (Bagdogra Airport Civil Enclave project), a major task involved annotating master AutoCAD structural plans with the total weight of every truss and column — data needed by the erection team for tower crane positioning and lift planning.

Each structural element's weight had to be read off a multi-row component material list table embedded in its individual fabrication drawing. With ~1500 drawings to get through, manually reading and transcribing each table was slow and error-prone.

This was **not a formally assigned task** — it was built independently to remove a bottleneck that was slowing down the broader documentation effort.

## What it does

1. Takes a folder of fabrication drawing screenshots (each containing a component material list table: mark, profile, material, number of pieces, length, area, weight).
2. Passes each screenshot to a **Qwen2.5-VL** vision-language model, running locally via **Ollama**, with a structured prompt asking for the table data back as JSON.
3. Parses the JSON response for each drawing.
4. Writes the extracted rows into a structured Excel workbook using `openpyxl` — ready for cross-referencing against the master AutoCAD plans.

```
Screenshots (drawings) --> Qwen2.5-VL (via Ollama) --> JSON --> Excel (openpyxl)
```

## Tech stack

- **Python** — orchestration, JSON parsing, Excel writing
- **Ollama** — local inference server for running the VLM without sending drawings to an external API
- **Qwen2.5-VL** — vision-language model used for reading the table images
- **openpyxl** — structured Excel output

## Why a local VLM instead of an API

Fabrication drawings for a live infrastructure project are not something you want leaving the local network. Running Qwen2.5-VL locally via Ollama kept all drawing data on-site while still getting VLM-quality table extraction.

## Usage

```bash
# Install dependencies
pip install -r requirements.txt

# Pull the model (one-time)
ollama pull qwen2.5vl

# Run the extraction
python extract_tables.py --input ./screenshots --output material_list.xlsx --model qwen2.5vl:latest
```

## Sample output

Each row extracted corresponds to one structural member:

| Mark | Profile | Material | No. | Length (mm) | Area (m²) | Weight (kg) |
|------|---------|----------|-----|-------------|-----------|-------------|
| BC102 | RHS300×200×12 | E350 | 1 | 15949 | 15.8 | 1430.3 |
| TR6 | RHS300×200×12 | E350 | 1 | 14748 | 14.6 | 1322.6 |
| M28 | SHS132×8 | E350 | 4 | 4012 | 2.1 | 486.6 |

*(Sample values shown for format illustration — actual project drawings are not included in this repo for confidentiality reasons.)*

## Result

Processed ~1500 fabrication drawings, producing a structured dataset that fed directly into the master AutoCAD weight annotations used by the erection team for tower crane positioning and lift planning — replacing what would otherwise have been fully manual transcription.

## Notes

- No project-specific drawings, client data, or proprietary information are included in this repository — only the automation code and illustrative/sample data.
- Built and used during a BITS Pilani Practice School-I internship at Kalpataru Projects International Ltd.
