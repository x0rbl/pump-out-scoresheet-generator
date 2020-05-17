# pump-out-scoresheet-generator

This script extracts data from the [Pump Out](https://pumpout.anyhowstep.com/) project and organizes it into an XLSX spreadsheet, which can be used for score tracking.

To use this script, you will need:

1. [Python 3](https://www.python.org/downloads/)
1. The [openpyxl](https://openpyxl.readthedocs.io/en/stable/) library (`python3 -m pip install openpyxl`)
1. A copy of the latest Pump Out database from, available [here](https://github.com/AnyhowStep/pump-out-sqlite3-dump/tree/master/dump)

Then do the following:

1. Copy the database to the directory with the script
1. Make sure the `DBPATH = '...'` line at the top of `generate.py` matches the filename of the database dump
1. Open `config.txt` and follow the (short) instructions to set the options; save the file when done
1. Run `python3 generate.py`
1. You should now have a new score sheet named `output.xlsx`

To quickly find the rows you want to update, you can use the filtering features of your spreadsheet editor to only show the modes and difficulties that you are interested in.