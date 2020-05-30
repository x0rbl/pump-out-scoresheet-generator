# pump-out-scoresheet-generator

This script extracts data from the [Pump Out](https://pumpout.anyhowstep.com/) project and organizes it into an XLSX spreadsheet, which can be used for score tracking.

## Dependencies

To use this script, you will need:

1. [Python 3](https://www.python.org/downloads/)
1. The [openpyxl](https://openpyxl.readthedocs.io/en/stable/) library (`python3 -m pip install openpyxl`)
1. A copy of the latest Pump Out database, available [here](https://github.com/AnyhowStep/pump-out-sqlite3-dump/tree/master/dump)

## Simple instructions

To create a new score sheet (`scores.xlsx`):

1. Open `config.txt` and follow the instructions to set the options; save the file when done
1. Run `python3 generate.py <path of database> scores.xlsx`

Later on, when the Pump Out database is updated with new songs or bugfixes, you will want to regenerate your score sheet.  Use the following to create a new score sheet (`newscores.xlsx`) pre-populated with your existing scores:

`python3 generate.py <path of database> newscores.xlsx --from scores.xlsx`

## Command-line options

`python3 <db> <out> [--from <from>] [--config <config>] [--overwrite]`

* `db`: The path of the Pump Out database
* `out`: The path to write the spreadsheet to
* `from`: The path of another spreadsheet containing
* `config`: The path of the configuration file (defalts to `config.txt`)
* `overwrite`: If specified, allow an existing score sheet to be overwritten

## Configuration options

`config.txt` allows you to configure the following aspects of the generated spreadsheet:

* Which mixes to include
* Which modes to include
* Which difficulties to include
* Whether to track scores for pad, keyboard, or both
* Whether to sort difficulties high-to-low or low-to-high
