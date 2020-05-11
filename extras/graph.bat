@echo off

:: This requires GraphViz to be installed
:: https://graphviz.gitlab.io/_pages/Download/Download_windows.html

dot db_relations.txt -Tpng -orelations.png && start mspaint relations.png
