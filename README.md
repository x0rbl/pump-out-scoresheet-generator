# pump-out-scoresheet-generator

This script extracts data from the [Pump Out](https://pumpout.anyhowstep.com/) project and organizes it into an XLSX spreadsheet, which can be used for score tracking.

To use this script, you will need:

1. [Python 3](https://www.python.org/downloads/)
1. The [openpyxl](https://openpyxl.readthedocs.io/en/stable/) library (`python3 -m pip install openpyxl`)
1. A copy of the latest Pump Out database from, available [here](https://github.com/AnyhowStep/pump-out-sqlite3-dump/tree/master/dump)

Then run `python3 generate.py`

To quickly find the rows you want to update, you can use the filtering features of your spreadsheet editor to only show the modes and difficulties that you are interested in.