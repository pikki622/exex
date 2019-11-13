# exex [![Build Status](https://travis-ci.org/vikpe/python-package-starter.svg?branch=master)](https://travis-ci.org/vikpe/python-package-starter) [![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
> Extract data from Excel documents

## Features
* Extract data from Excel (xlsx)
* Format result as JSON, JSONL, XML

## Installation
```sh
pip install exex
```

## Usage

![Sample Excel file](https://raw.githubusercontent.com/vikpe/exex/master/docs/sample_xlsx.png "Sample Excel file")

```python
from exex import extract

ext = extract.Extractor('sample.xlsx')

ext.sheets().active()
ext.sheets().first()
ext.sheets().last()
ext.sheets().nth()
ext.sheet("prices").cells().all().json()

ext.all()

ext.range("A1:B2")
ext.range("A1:A4")
ext.ranges(["A1:B2", "A1:A4"])

ext.cell("A1")
ext.cells(["A1", "B2"])

ext.column("A")
ext.columns(["A", "B", "C"])
ext.column_range(from="C", to="D")

ext.row(1)
ext.rows([1,2,3])
ext.row_range(from=3, to=5)
```

## Development

**Tests** (local Python version)
```sh
poetry run pytest
```

**Tests** (all Python versions defined in `tox.ini`)
```sh
poetry run tox
```

**Code formatting** (black)
```sh
poetry run black .
```
