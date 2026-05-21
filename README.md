# OCI Cost Updater

Python tool to fetch monthly Oracle Cloud Infrastructure (OCI) costs grouped by compartment and write them into an existing Excel workbook.

The project is designed for a simple monthly run, but avoids hardcoded production values. Local settings live in `.env`, credentials remain in the standard OCI SDK config, and command-line arguments can override monthly values when needed.

## Features

- Fetches OCI cost usage grouped by `compartmentName`
- Supports subcompartments through configurable `compartment_depth`
- Updates an existing Excel worksheet by matching compartment names
- Writes `0.0` for worksheet compartments without OCI cost in the period
- Reports OCI compartments that were not found in the worksheet
- Supports `--dry-run` to validate without saving the workbook
- Returns non-zero exit codes for configuration and runtime failures

## Requirements

- Python 3.10 or newer
- OCI SDK credentials configured in `~/.oci/config`
- Read access to OCI usage/cost data
- An Excel workbook with compartment names in a known column

## Installation

```bash
git clone https://github.com/Celta031/OCI-Cost-Automation.git
cd OCI-Cost-Automation

python -m venv venv
venv\Scripts\activate

pip install -r requirements.txt
```

For development and tests:

```bash
pip install -r requirements-dev.txt
```

## Configuration

Copy the example file and adjust the values for your monthly run:

```bash
copy .env.example .env
```

Example `.env`:

```env
OCI_EXCEL_PATH=sua-planilha.xlsx
OCI_SHEET_NAME=OCI CONSUMO
OCI_START_ROW=5
OCI_COMPARTMENT_COLUMN=A
OCI_TARGET_COLUMN=S
OCI_BILLING_START=2026-04-21T00:00:00Z
OCI_BILLING_END=2026-05-21T00:00:00Z
OCI_PROFILE=DEFAULT
OCI_COMPARTMENT_DEPTH=6
```

Do not commit your real `.env`, OCI config, private keys, or production `.xlsx` files.

## Usage

Run with values from `.env`:

```bash
python oci_cost_updater.py
```

Override values from the command line:

```bash
python oci_cost_updater.py --excel-path sua-planilha.xlsx --sheet "OCI CONSUMO" --start 2026-04-21T00:00:00Z --end 2026-05-21T00:00:00Z --compartment-column A --target-column S --start-row 5
```

Validate without saving the workbook:

```bash
python oci_cost_updater.py --dry-run
```

Use a specific OCI profile:

```bash
python oci_cost_updater.py --oci-profile PROD
```

## CLI Options

- `--excel-path`: Excel workbook path
- `--sheet`: worksheet name; if omitted, the active sheet is used
- `--start-row`: first row containing compartment data
- `--compartment-column`: column with compartment names
- `--target-column`: column where costs are written
- `--start`: billing start in `YYYY-MM-DDTHH:MM:SSZ`
- `--end`: billing end in `YYYY-MM-DDTHH:MM:SSZ`
- `--oci-profile`: optional OCI config profile
- `--compartment-depth`: OCI compartment depth, default `6`
- `--env-file`: path to env file, default `.env`
- `--dry-run`: runs without saving the workbook
- `--log-level`: `DEBUG`, `INFO`, `WARNING`, or `ERROR`

Command-line values take precedence over environment variables and `.env`.

## Exit Codes

- `0`: success
- `1`: runtime failure, such as OCI or workbook save errors
- `2`: invalid configuration, such as missing file, invalid date, invalid column, or missing worksheet

## Tests

```bash
pytest
```

The automated tests cover config precedence, date validation, cost aggregation, worksheet updates, `dry-run`, and missing worksheet handling.

## Design

See [SDD.md](SDD.md) for the production design and implementation decisions.

## Local Legacy Script

`atualizar_custos_oci.py` is treated as a local legacy script and is not required for the public workflow. New monthly runs should use `oci_cost_updater.py`.
