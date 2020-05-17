from parse_pump_out import * # Clean up later
from parse_config import parse_config, verify_config

from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import CellIsRule, FormulaRule, Rule
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter

import datetime
import sys

###### PUT THE PATH OF YOUR DATABASE HERE #####
DBPATH = 'pumpout-2020-05-03-22-18-1588558611155.db'
###### PUT THE PATH OF YOUR DATABASE HERE #####

def adjust_column_widths(ws, cols, rows):
	for c in cols:
		width = 0
		for r in rows:
			val = ws.cell(row=r, column=c).value
			if val == None: continue
			width = max(width, len(str(val)))
		ws.column_dimensions[get_column_letter(c)].width = width + 1

def write_data_sheet(ws, rows, mixes):
	headers = [
		"CID", # A
		"SID", # B
		"GID", # C
		"Title", # D
		"Cut", # E
		"Mode", # F
		"Difficulty", # G
		"First Seen", # H
		"Last Seen", # I
		"BPM", # J
		"Category", # K
		"Stepmaker", # L
	]
	headers += mixes
	headers += [
		"Labels",
		"Card",
		"Comment"
	]
	m = len(mixes)
	for i in range(len(headers)):
		ws.cell(row=1, column=i+1, value=headers[i]).font = Font(bold=True)
	for i in range(len(rows)):
		first_seen = last_seen = "???"
		if rows[i].first_seen != None:
			first_seen = str(rows[i].first_seen)
		if rows[i].last_seen != None:
			last_seen = str(rows[i].last_seen)
		ws.cell(row=i+2, column=1, value=rows[i].cid)
		ws.cell(row=i+2, column=2, value=rows[i].sid)
		ws.cell(row=i+2, column=3, value=rows[i].gameIdentifier)
		ws.cell(row=i+2, column=4, value=rows[i].title)
		ws.cell(row=i+2, column=5, value=str(rows[i].cut))
		ws.cell(row=i+2, column=6, value=rows[i].rating.mode_full())
		ws.cell(row=i+2, column=7, value=rows[i].rating.difficulty_str())
		ws.cell(row=i+2, column=8, value=first_seen)
		ws.cell(row=i+2, column=9, value=last_seen)
		ws.cell(row=i+2, column=10, value=str(rows[i].bpm))
		ws.cell(row=i+2, column=11, value=rows[i].category)
		ws.cell(row=i+2, column=12, value=str(rows[i].stepmaker))
		for j in range(m):
			ws.cell(row=i+2, column=13+j, value="NY"[set([mixes[j]]) & rows[i].mixes != set()])
		ws.cell(row=i+2, column=13+m, value=",".join(rows[i].labels))
		ws.cell(row=i+2, column=14+m, value=rows[i].card)
		ws.cell(row=i+2, column=15+m, value=rows[i].comment)

	ws.column_dimensions['A'].width = 5
	ws.column_dimensions['B'].width = 5
	ws.column_dimensions['C'].width = 5
	ws.column_dimensions['D'].width = 36
	ws.column_dimensions['E'].width = 9
	ws.column_dimensions['F'].width = 9
	ws.column_dimensions['G'].width = 3
	ws.column_dimensions['H'].width = 17
	ws.column_dimensions['I'].width = 17
	ws.column_dimensions['J'].width = 8
	ws.column_dimensions['K'].width = 14
	ws.column_dimensions['L'].width = 31
	for i in range(m):
		ws.column_dimensions[get_column_letter(13+i)].width = 2
	ws.column_dimensions[get_column_letter(13+m)].width = 39
	ws.column_dimensions[get_column_letter(14+m)].width = 48
	ws.column_dimensions[get_column_letter(15+m)].width = 120

	ws.freeze_panes = 'A2'

def write_score_sheet(ws, rows, config):
	headers = [
		"CID",           # A
		"Title",         # B
		"Cut",           # C
		"Mode",          # D
		"Difficulty",    # E
		"Passed (Pad)",  # F
		"Grade (Pad)",   # G
		"Miss (Pad)",    # H
		"Comment (Pad)", # I
	]
	if config.keyboard:
		headers += [
			"Passed (Kbd)",  # J
			"Grade (Kbd)",   # K
			"Miss (Kbd)",    # L
			"Comment (Kbd)"  # M
		]
	bold = Font(bold=True)
	gray = PatternFill("solid", fgColor="EEEEEE")
	dgray = PatternFill("solid", fgColor="CCCCCC")
	bottom = Border(bottom=Side(style="thick"))
	right = Border(right=Side(style="thick"))
	for i in range(len(headers)):
		c = ws.cell(row=1, column=i+1, value=headers[i])
		c.font = bold
		c.fill = dgray
		c.border = bottom
	for i in range(len(rows)):
		ws.cell(row=i+2, column=1, value=rows[i].cid).fill = dgray
		ws.cell(row=i+2, column=2, value="=VLOOKUP(A%d, 'Data (Complete)'!A1:O9999, 4, FALSE)" % (i+2)).fill = gray
		ws.cell(row=i+2, column=3, value="=VLOOKUP(A%d, 'Data (Complete)'!A1:O9999, 5, FALSE)" % (i+2)).fill = gray
		ws.cell(row=i+2, column=4, value="=VLOOKUP(A%d, 'Data (Complete)'!A1:O9999, 6, FALSE)" % (i+2)).fill = gray
		ws.cell(row=i+2, column=5, value="=VLOOKUP(A%d, 'Data (Complete)'!A1:O9999, 7, FALSE)" % (i+2)).fill = gray
		ws.cell(row=i+2, column=5).border = right
		if config.keyboard:
			ws.cell(row=i+2, column=9).border = right

	for i in range(len(rows)):
		ws.cell(row=i+2, column=9).alignment = Alignment(horizontal='fill')

	thin = Border(bottom=Side(style="thin", color="888888"))
	for c in range(len(headers)):
		for r in range(len(rows)-1):
			ws.cell(row=r+2, column=c+1).border = thin

	ws.column_dimensions['A'].width = 5
	ws.column_dimensions['B'].width = 30
	ws.column_dimensions['C'].width = 9
	ws.column_dimensions['D'].width = 9
	ws.column_dimensions['E'].width = 3
	ws.column_dimensions['F'].width = 4
	ws.column_dimensions['G'].width = 4
	ws.column_dimensions['H'].width = 4
	ws.column_dimensions['I'].width = 16
	if config.keyboard:
		ws.column_dimensions['J'].width = 4
		ws.column_dimensions['K'].width = 4
		ws.column_dimensions['L'].width = 4
		ws.column_dimensions['M'].width = 16

	green = PatternFill("solid", bgColor="44FF44")
	red = PatternFill("solid", bgColor="FF4444")
	grade_f = PatternFill("solid", bgColor="555555")
	grade_d = PatternFill("solid", bgColor="666666")
	grade_c = PatternFill("solid", bgColor="777777")
	grade_b = PatternFill("solid", bgColor="888888")
	grade_a = PatternFill("solid", bgColor="AAAAAA")
	grade_s = PatternFill("solid", bgColor="DD88FF")
	grade_ss = PatternFill("solid", bgColor="FFFF00")
	grade_sss = PatternFill("solid", bgColor="44FF44")

	ws.conditional_formatting.add('F2:F9999', CellIsRule(operator='equal', formula=['"Y"'], fill=green))
	ws.conditional_formatting.add('F2:F9999', CellIsRule(operator='equal', formula=['"N"'], fill=red))

	if config.keyboard:
		ws.conditional_formatting.add('J2:J9999', CellIsRule(operator='equal', formula=['"Y"'], fill=green))
		ws.conditional_formatting.add('J2:J9999', CellIsRule(operator='equal', formula=['"N"'], fill=red))

	ws.conditional_formatting.add('G2:G9999', CellIsRule(operator='equal', formula=['"SSS"'], fill=grade_sss))
	ws.conditional_formatting.add('G2:G9999', CellIsRule(operator='equal', formula=['"SS"'], fill=grade_ss))
	ws.conditional_formatting.add('G2:G9999', CellIsRule(operator='equal', formula=['"S"'], fill=grade_s))
	ws.conditional_formatting.add('G2:G9999', CellIsRule(operator='equal', formula=['"A"'], fill=grade_a))
	ws.conditional_formatting.add('G2:G9999', CellIsRule(operator='equal', formula=['"B"'], fill=grade_b))
	ws.conditional_formatting.add('G2:G9999', CellIsRule(operator='equal', formula=['"C"'], fill=grade_c))
	ws.conditional_formatting.add('G2:G9999', CellIsRule(operator='equal', formula=['"D"'], fill=grade_d))
	ws.conditional_formatting.add('G2:G9999', CellIsRule(operator='equal', formula=['"F"'], fill=grade_f))

	if config.keyboard:
		ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"SSS"'], fill=grade_sss))
		ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"SS"'], fill=grade_ss))
		ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"S"'], fill=grade_s))
		ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"A"'], fill=grade_a))
		ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"B"'], fill=grade_b))
		ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"C"'], fill=grade_c))
		ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"D"'], fill=grade_d))
		ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"F"'], fill=grade_f))

	ws.freeze_panes = 'C2'

def write_rows(xlpath, rows, dbversion, config, mixes_all):
	wb = Workbook()
	ws_scores = wb.active
	ws_scores.title = "Scores"

	mf = set(config.mixes)
	rows_filter = []
	for r in rows:
		if not (r.mixes & mf):
			continue
		if not (r.rating.mode_full() in config.modes):
			continue
		d = r.rating.difficulty
		if d == None and not config.unrated:
			continue
		if d != None and (d < config.diff_min or d > config.diff_max):
			continue

		rows_filter.append(r)

	write_score_sheet(ws_scores, rows_filter, config)

	ws_dump = wb.create_sheet(title="Data")
	write_data_sheet(ws_dump, rows_filter, config.mixes)

	ws_dump = wb.create_sheet(title="Data (Complete)")
	write_data_sheet(ws_dump, rows, mixes_all)

	bold = Font(bold=True)
	ws_marker = wb.create_sheet(title="About")
	ws_marker.cell(row=1, column=1, value="Database Name:").font = bold
	ws_marker.cell(row=1, column=2, value=dbversion)
	ws_marker.cell(row=2, column=1, value="Sheet Generated On:").font = bold
	ws_marker.cell(row=2, column=2, value=str(datetime.datetime.now()))
	ws_marker.cell(row=3, column=1, value="Generator Version:").font = bold
	ws_marker.cell(row=3, column=2, value="v0.4")

	ws_marker.cell(row=5, column=1, value="Player:").font = bold
	ws_marker.cell(row=5, column=2, value="[YOUR NAME HERE]").font = bold
	ws_marker.cell(row=6, column=1, value="Mixes:").font = bold
	ws_marker.cell(row=6, column=2, value=", ".join(config.mixes))
	ws_marker.cell(row=7, column=1, value="Modes:").font = bold
	ws_marker.cell(row=7, column=2, value=", ".join(config.modes))
	ws_marker.cell(row=8, column=1, value="Difficulties:").font = bold
	ws_marker.cell(row=8, column=2, value="%d-%d%s" % (config.diff_min, config.diff_max, (""," (+Unrated)")[config.unrated]))
	ws_marker.cell(row=9, column=1, value="Options:").font = bold
	ws_marker.cell(row=9, column=2, value="%s" % ("","+Keyboard")[config.keyboard])
	
	adjust_column_widths(ws_marker, range(1,2+1), range(1,8+1))
	
	wb.save(xlpath)
	wb.close()

def generate_xlsx(dbpath, xlpath, config):
	db = read_database(dbpath)
	rows = flatten_db(db)
	mixes_all = list(db.mixes.values())
	mixes_all.sort(key=lambda e: e.sortOrder)
	mixes_all = [m.title for m in mixes_all]
	verify_config(config, mixes_all, db.modes)
	write_rows(xlpath, rows, dbpath, config, mixes_all)

if __name__ == '__main__':
	config = parse_config("config.txt")
	generate_xlsx(DBPATH, "output.xlsx", config)
