## Seaward Super String Parser

### Overview

The Seaward Super String Parser is a Python tool for parsing Seaward .sss test files and exporting the test data into an
easily readable Excel file.

### Requirements

Install dependencies:

```bash
pip install -r requirements.txt
```

### Usage

Run from the command line:

```bash
python parser.py <sss_file_path>
```

* <sss_file_path>: Path to the .sss file you want to parse

* Output Excel file will be saved to current working directory

An example file testResults.sss is provided for quick testing, example:

```bash
python parser.py testResults.sss
```