# TerraBrasilis/INPE Public Data Downloader

This repository provides a simple tool to list, download, resume interrupted downloads, and validate public ZIP files available from the official TerraBrasilis/INPE download page.

Data source:

https://terrabrasilis.dpi.inpe.br/en/download-files/

This repository provides code only. It does not redistribute TerraBrasilis/INPE datasets. All files are downloaded directly from the official source and stored locally on the user's computer.

## Repository structure

```text
terrabrasilis-inpe-downloader/
├── README.md
├── LICENSE
├── requirements.txt
├── .gitignore
├── CITATION.cff
├── scripts/
│   └── download_terrabrasilis.py
└── data/
    ├── raw/
    └── metadata/
```

## Installation

Create a virtual environment and install the required packages:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

On Linux or macOS:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Basic use

Run the script with manual confirmation before downloading files:

```bash
python scripts/download_terrabrasilis.py
```

Run the script without interactive confirmation:

```bash
python scripts/download_terrabrasilis.py --yes
```

Set a local output folder:

```bash
python scripts/download_terrabrasilis.py --root data/raw/terrabrasilis --yes
```

Force dynamic page rendering with Selenium, in case the download links are no longer available in the static HTML:

```bash
python scripts/download_terrabrasilis.py --dynamic --yes
```

## Local outputs

By default, downloaded files are saved to:

```text
data/raw/terrabrasilis/<YYYY-MM-DD>/
```

Inside this folder, the script organizes files by biome and category when this information can be inferred from the file URL.

The script also saves metadata and validation files:

```text
terrabrasilis_zips.csv
terrabrasilis_zips.json
validation_report.json
```

These local files should not be committed to GitHub.

## Data use notice

This repository does not redistribute TerraBrasilis/INPE datasets. The scripts only automate downloads directly from the official source.

Users are responsible for checking and following the current terms of use, citation requirements, and data-use conditions defined by INPE/TerraBrasilis before publishing, sharing, or interpreting any derived products.

Outputs derived from these data should be validated before use in reports, publications, dashboards, technical analyses, or decision-making processes.

## License

The code in this repository is released under the MIT License.

This license applies only to the code developed in this repository. It does not apply to TerraBrasilis/INPE datasets or to any other data downloaded by the scripts.
