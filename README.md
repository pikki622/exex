# exex [![Build Status](https://travis-ci.org/vikpe/python-package-starter.svg?branch=master)](https://travis-ci.org/vikpe/python-package-starter) [![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
> Excel extractor written in Python

## Features
* Extract data from Excel (xlsx)
* Format result as JSON, JSONL, XML

## Installation
```sh
pip install exex
```

## Usage
```python
from exex import Extractor

ext = Extractor('sample.xlsx')
ext.all()
ext.range("A1:B2")
ext.cell("A1")
ext.cells("A1", "B2")
```

Command | Description
--- | ---
`todo` | TODO

### API
TODO

## Development
TODO
