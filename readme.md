# Likert MSI Template Calculator

A Python-based pipeline for calculating the Midpoint Score Index (MSI) from Likert-scale questionnaire data. This repository automates data processing by reading responses from an Excel file, applying optional reverse coding, computing statistical metrics (frequency, cumulative proportions, z-scores, and corrected z-scores), and exporting a consolidated report.

---

## Table of Contents

1. [Features](#features)
2. [Prerequisites](#prerequisites)
3. [Installation](#installation)
4. [Configuration](#configuration)
5. [Input Data Format](#input-data-format)
6. [Usage](#usage)
7. [Output Files](#output-files)
8. [Script Overview](#script-overview)
9. [Troubleshooting](#troubleshooting)
10. [License](#license)

---

## Features

- **Automatic metric computation** for each survey item:
  - **F**: Frequency count (1–5)
  - **P**: Proportion (F / N)
  - **CP**: Cumulative proportion
  - **MID\_CP**: Midpoint cumulative (CP − P/2)
  - **0.5−MID\_CP**
  - **Z**: Clipped z-score (±3.9 bounds)
  - **ZC**: Zero-based corrected z-score (Z − min(Z))
  - **Rounded scores** (integer)
- **Reverse coding** support via environment-configurable list of item codes.
- **Dynamic item detection**: no hard-coded question list—reads all columns in input.
- **Configurable** via a `.env` file (input/output paths, sheet name, reverse-coded items).
- **Progress logging** to terminal for transparency.

---

## Prerequisites

- **Python** 3.7 or higher
- **pip** (Python package manager)

### Required Python Packages

Install the dependencies:

```bash
pip install pandas scipy python-dotenv xlsxwriter openpyxl
```

---

## Installation

Clone this repository and navigate into its directory:

```bash
git clone https://github.com/ishadiyantos/likert_and_transpose.git
cd likert_and_transpose
```

Install the Python dependencies as shown above.

---

## Configuration

Create a file named `.env` in the repository root with the following content:

```dotenv
# Path to the Excel file containing raw responses\
INPUT_FILE=responses.xlsx

# Name of the worksheet with response data\
INPUT_SHEET=Sheet1

# Desired output filename for the MSI report\
OUTPUT_FILE=MSI_all_in_one.xlsx

# (Optional) Comma-separated list of column headers to reverse-code\
# Example: REVERSE_ITEMS=Q3,Q7,Q12
REVERSE_ITEMS=
```

- **INPUT\_FILE**: Excel file in the project directory.
- **INPUT\_SHEET**: Worksheet name with headers matching survey question codes.
- **OUTPUT\_FILE**: Name for the generated report.
- **REVERSE\_ITEMS**: If you have reverse-scored items, list their column codes here.

---

## Input Data Format

The input Excel must contain a header row of question codes and subsequent rows of integer responses (1–5). Optionally, you may include a respondent ID column, which will be ignored by the script.

Example:

| respondent\_id | Q1  | Q2  | Q3  | ... |
| -------------- | --- | --- | --- | --- |
| 1              | 4   | 2   | 5   | ... |
| 2              | 3   | 1   | 4   | ... |
| ...            | ... | ... | ... | ... |

---

## Usage

Run the main script from the command line:

```bash
python3 msi_template.py
```

You will see logs indicating configuration, data reading, item processing, and file writing:

```
[msi_template] Config → INPUT_FILE='responses.xlsx', INPUT_SHEET='Sheet1', OUTPUT_FILE='MSI_all_in_one.xlsx'
[msi_template] Reverse-coded items: (none)
[msi_template] Reading data ...
[msi_template] Data shape: 150 rows × 80 columns
[msi_template] Processing metrics for 80 items
[msi_template] Sheet1 built: 720 rows
[msi_template] Sheet2 built: 150 rows × 80 columns
[msi_template] Writing output to 'MSI_all_in_one.xlsx'
[msi_template] Done.
```

---

## Output Files

The script outputs a single Excel file (`OUTPUT_FILE`):

- **Sheet1**: Detailed metric blocks for each question code, with two blank rows between blocks.
- **Sheet2**: Respondent-level table of rounded corrected z-scores for each question.

---

## Script Overview

- ``: Main entry point. Loads configuration, reads data, computes metrics, writes output.
- ``: Returns frequency, proportions, cumulative sums, MID\_CP, z-scores, corrected z-scores, and rounded values.
- **Reverse coding**: Values transformed via `6 − x` if the item code appears in `REVERSE_ITEMS`.
- **Performance**: Avoids DataFrame fragmentation by constructing `Sheet2` in a single dictionary-to-DataFrame step.

---

## Troubleshooting

- **FileNotFoundError**: Verify `INPUT_FILE` and `INPUT_SHEET` in `.env` match existing files.
- **Empty output**: Ensure your sheet contains numeric responses 1–5 and column headers are correct.
- **Missing dependencies**: Make sure to install all required Python packages.

---

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

