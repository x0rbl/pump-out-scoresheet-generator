from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import CellIsRule, FormulaRule, Rule
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.styles.differential import DifferentialStyle

import bisect
import datetime
import sqlite3
import sys

###### PUT THE PATH OF YOUR DATABASE HERE #####
DBPATH = 'pumpout-2020-05-03-22-18-1588558611155.db'
###### PUT THE PATH OF YOUR DATABASE HERE #####

class Version:
	def __init__(self, versionId, mixTitle, versionTitle, parentId, sortOrder):
		self.versionId = versionId
		self.mixTitle = mixTitle
		self.versionTitle = versionTitle
		self.parentId = parentId
		self.sortOrder = sortOrder

	def __str__(self):
		name = self.mixTitle
		if self.versionTitle != "default":
			name += " " + self.versionTitle
		return name

class VersionedValue:
	def __init__(self, versions={}):
		self.values = {}
		self.cache = None
		self.add_versions(versions)

	def add_versions(self, versions):
		self.versions = versions

	def add(self, versionId, value):
		if versionId in self.values:
			raise Exception("duplicate key: vid=%d, new=%s, old=%s" % (versionId, value, self.values[versionId]))
		self.values[versionId] = value
		self.cache = None

	def recent(self, versionId, default):
		while True:
			if not versionId in self.versions:
				return default
			if not versionId in self.values:
				versionId = self.versions[versionId].parentId
				continue
			return self.values[versionId]

	def at(self, versionId, default):
		return self.at_version(versionId, default)[1]

	def at_version(self, versionId, default):
		if not versionId in self.versions:
			return default
		so = self.versions[versionId].sortOrder
		self.ensure_cache()
		for v in self.cache:
			if v.sortOrder <= so:
				return (v, self.values[v.versionId])
		return (None, default)

	def latest(self, default):
		if len(self.values) == 0:
			return default
		self.ensure_cache()
		return self.values[self.cache[0].versionId]

	def ensure_cache(self):
		if self.cache == None:
			self.cache = sorted([self.versions[vid] for vid in self.values.keys()], key=lambda e: -e.sortOrder)

class MultipleVersionedValue:
	def __init__(self):
		self.versions = {}
		self.values = {}

	def add_versions(self, versions):
		self.versions = versions
		for vv in self.values.values():
			vv.add_versions(versions)

	def add(self, versionId, operation, value):
		if not value in self.values:
			self.values[value] = VersionedValue(self.versions)
		self.values[value].add(versionId, operation)

	def all(self, version):
		res = []
		for value, vv in self.values.items():
			if vv.at(version, "DELETE") != "DELETE":
				res.append(value)
		return res

	def one(self, version, default):
		vals = self.all(version)
		if len(vals) > 1:
			raise Exception("too many vals: %s" % vals)
		if vals == []:
			return default
		return vals[0]

	def best(self, version, default):
		res = []
		for value, vv in self.values.items():
			(atver, atop) = vv.at_version(version, "DELETE")
			if vv.at_version(version, "DELETE")[1] != "DELETE":
				res.append((atver, value))

		if len(res) == 0:
			return default
		temp = [(ver.sortOrder, val) for (ver,val) in res]
		temp.sort(reverse=True)
		return temp[0][1]

class Rating:
	def __init__(self, mode='?', difficulty=None, path=None):
		self.mode = mode
		self.difficulty = difficulty
		self.path = path

	def order(self, down=True):
		seq = ['S', 'SP', 'D', 'DP', 'HDB', 'C', 'R']

		if self.difficulty == None:
			return (len(seq) + 1) ** 2

		m = len(seq)
		if self.mode in seq:
			m = seq.index(self.mode)

		d = self.difficulty * (len(seq) + 1)
		if down:
			d = -d

		return d + m

	def mode_full(self):
		fulls = {"S":"Single", "D":"Double", "SP":"Single Performance", "DP":"Double Performance", "HDB":"Half-Double", "C":"Co-Op", "R":"Routine"}
		return fulls.get(self.mode, self.mode)

	def difficulty_str(self):
		if self.difficulty == None:
			return "??"
		return "%02d" % self.difficulty

	def __str__(self):
		if self.difficulty != None:
			return "%s%02d" % (self.mode, self.difficulty)
		return "%s??" % self.mode

class Bpm:
	def __init__(self, low=-1.0, high=-1.0):
		self.low = low
		self.high = high

	def __str__(self):
		if self.low == self.high:
			if int(self.low) == self.low:
				return "%d" % int(self.low)
			return "%s" % self.low
		if int(self.low) == self.low and int(self.high) == self.high:
			return "%d-%d" % (self.low, self.high)
		return "%s-%s" % (self.low, self.high)

class Stepmaker:
	def __init__(self):
		self.str = ""
		self.list = []

	def add(self, prefix, name):
		if self.str != "" and prefix != "":
			self.str += " "
		self.str += prefix

		if self.str != "" and name != "":
			self.str += " "
		self.str += name

		self.list.append(name)

	def __str__(self):
		return self.str

class Labels:
	def __init__(self):
		self.versions = {}
		self.labels = {}

	def add_versions(self, versions):
		self.versions = versions
		for vv in self.labels.values():
			vv.add_versions(versions)

	def add(self, versionId, operation, label):
		if not label in self.labels:
			self.labels[label] = VersionedValue(self.versions)
		self.labels[label].add(versionId, operation)

	def at(self, version):
		res = []
		for label, vv in self.labels.items():
			if vv.at(version, "DELETE") != "DELETE":
				res.append(label)
		return res

class Cut:
	def __init__(self, cut="NOCUT"):
		self.cut = cut

	def order(self):
		seq = ['Short Cut','Arcade','Remix','Full Song']
		if self.cut in seq:
			return seq.index(self.cut)
		return len(seq) + 1

	def __str__(self):
		return self.cut

class Song:
	def __init__(self):
		self.songId = -1
		self.operations = VersionedValue()
		self.comment = VersionedValue()
		self.title = VersionedValue()
		self.gameIdentifier = MultipleVersionedValue()
		self.category = VersionedValue()
		self.bpm = VersionedValue()
		self.card = MultipleVersionedValue()
		self.cut = Cut()
		self.fallbackTitle = "NOTITLE"

	def add_versions(self, versions):
		self.operations.add_versions(versions)
		self.comment.add_versions(versions)
		self.title.add_versions(versions)
		self.gameIdentifier.add_versions(versions)
		self.category.add_versions(versions)
		self.bpm.add_versions(versions)
		self.card.add_versions(versions)

class Chart:
	def __init__(self):
		self.chartId = -1
		self.songId = -1
		self.operations = VersionedValue()
		self.rating = VersionedValue()
		self.stepmaker = Stepmaker()
		self.labels = MultipleVersionedValue()

	def add_versions(self, versions):
		self.operations.add_versions(versions)
		self.rating.add_versions(versions)
		self.labels.add_versions(versions)

class Entry:
	def __init__(self):
		self.cid = -1
		self.sid = -1
		self.name = "NONAME"
		self.rating = "NORATING"
		self.stepmaker = Stepmaker()
		self.labels = []
		self.comment = ""
		self.cut = "NOCUT"
		self.gameIdentifier = "NOID"
		self.category = "NOCATEGORY"
		self.bpm = Bpm()
		self.card = "NOCARD"
		self.first_seen = None
		self.last_seen = None

	def __str__(self):
		return "[cid=%d, sid=%d, name=%s, rating=%s, stepmaker=%s, labels=%s, comment=%s, cut=%s, gameIdentifier=%s, category=%s, bpm=%s, card=%s]" % (self.cid, self.sid, self.title, self.rating, self.stepmaker, self.labels, self.comment, self.cut, self.gameIdentifier, self.category, self.bpm, self.card)

def latest_version(versions):
	if len(versions) == 0:
		raise Exception("no versions")
	latest = None
	so = -1
	for v in versions.values():
		if v.sortOrder > so:
			so = v.sortOrder
			latest = v
	return latest

def read_database(dbpath):
	conn = sqlite3.connect(dbpath)
	c = conn.cursor()

	### Get version/mix tree
	c.execute('''
		SELECT
			version.versionId,
			mix.internalTitle,
			version.internalTitle,
			version.parentVersionId,
			version.sortOrder
		FROM version
		JOIN mix ON version.mixId = mix.mixId
	''')
	versions = {}
	for versionId, mixTitle, versionTitle, parentId, sortOrder in c.fetchall():
		versions[versionId] = Version(versionId, mixTitle, versionTitle, parentId, sortOrder)

	### Populate charts by version
	c.execute('''
		SELECT
			chartVersion.chartId,
			chart.songId,
			chartVersion.versionId,
			operation.internalTitle
		FROM chartVersion
		JOIN chart ON chart.chartId = chartVersion.chartId
		JOIN operation ON chartVersion.operationId = operation.operationId
	''')
	charts = {}
	for chartId, songId, versionId, operation in c.fetchall():
		if not chartId in charts:
			ch = Chart()
			ch.chartId = chartId
			ch.songId = songId
			ch.add_versions(versions)
			charts[chartId] = ch
		charts[chartId].operations.add(versionId, operation)

	### Get chart ratings
	c.execute('''
		SELECT
			chartRatingVersion.chartId,
			chartRatingversion.versionId,
			mode.internalAbbreviation,
			difficulty.value,
			rating.path
		FROM chartRatingVersion
		JOIN chartRating on chartRatingVersion.chartRatingId = chartRating.chartRatingId
		JOIN difficulty ON chartRating.difficultyId = difficulty.difficultyId
		JOIN mode ON chartRating.modeId = mode.modeId
		JOIN rating ON chartRating.modeId = rating.modeId AND chartRating.difficultyId = rating.difficultyId
	''')
	for chartId, versionId, mode, difficulty, path in c.fetchall():
		charts[chartId].rating.add(versionId, Rating(mode, difficulty, path))

	### Get chart stepmakers
	c.execute('''
		SELECT
			chartStepmaker.chartId,
			chartStepmaker.prefix,
			stepmaker.internalTitle,
			chartStepmaker.sortOrder
		FROM chartStepmaker
		JOIN stepmaker ON chartStepmaker.stepmakerId = stepmaker.stepmakerId
		ORDER BY chartId ASC, chartStepmaker.sortOrder ASC
	''')
	for chartId, prefix, stepmaker, _ in c.fetchall():
		charts[chartId].stepmaker.add(prefix, stepmaker)

	### Get chart labels
	c.execute('''
		SELECT
			chartLabel.chartId,
			chartLabelVersion.versionId,
			operation.internalTitle,
			label.internalTitle
		FROM chartLabelVersion
		JOIN chartLabel ON chartLabelVersion.chartLabelId = chartLabel.chartLabelId
		JOIN label ON chartLabel.labelId = label.labelId
		JOIN operation ON chartLabelVersion.operationId = operation.operationId
	''')
	for chartId, versionId, operation, label in c.fetchall():
		charts[chartId].labels.add(versionId, operation, label)

	### Create songs by version, with cut (Full Song, Remix, etc) and fallback title
	c.execute('''
		SELECT
			songVersion.songId,
			songVersion.versionId,
			operation.internalTitle,
			songVersion.internalDescription,
			cut.internalTitle,
			song.internalTitle
		FROM songVersion
		JOIN song ON songVersion.songId = song.songId
		JOIN operation ON songVersion.operationId = operation.operationId
		JOIN cut ON song.cutId = cut.cutId
	''')
	songs = {}
	for songId, versionId, operation, comment, cut, fallbackTitle in c.fetchall():
		if not songId in songs:
			songs[songId] = Song()
			songs[songId].songId = songId
			songs[songId].add_versions(versions)
		songs[songId].comment.add(versionId, (operation, comment))
		songs[songId].cut = Cut(cut)
		songs[songId].fallbackTitle = fallbackTitle

	### Get song titles
	c.execute('''
		SELECT
			songTitleVersion.songId,
			songTitleVersion.versionId,
			songTitle.title
		FROM songTitleVersion
		JOIN songTitle ON songTitleVersion.songTitleId = songTitle.songTitleId
		AND songTitleVersion.languageId = songTitle.languageId
		JOIN language ON songTitle.languageId = language.languageId
		WHERE language.code = "en"
	''')
	for songId, versionId, title in c.fetchall():
		songs[songId].title.add(versionId, title)

	### Get official song identifiers
	c.execute('''
		SELECT
			songGameIdentifier.songId,
			songGameIdentifier.gameIdentifier,
			songGameIdentifierVersion.versionId,
			operation.internalTitle
		FROM songGameIdentifierVersion
		JOIN songGameIdentifier ON songGameIdentifierVersion.songGameIdentifierId = songGameIdentifier.songGameIdentifierId
		JOIN operation ON songGameIdentifierVersion.operationId = operation.operationId
	''')
	for songId, gameIdentifier, versionId, operation in c.fetchall():
		songs[songId].gameIdentifier.add(versionId, operation, gameIdentifier)

	### Get song categories (K-Pop, World Music, etc)
	c.execute('''
		SELECT
			songCategoryVersion.songId,
			category.internalTitle,
			songCategoryVersion.versionId
		FROM songCategoryVersion
		JOIN songCategory ON songCategoryVersion.songCategoryId = songCategory.songCategoryId
		JOIN category ON songCategory.categoryId = category.categoryId
	''')
	for songId, category, versionId in c.fetchall():
		songs[songId].category.add(versionId, category)

	### Get song BPM info
	c.execute('''
		SELECT
			songBpmVersion.songId,
			songBpmVersion.versionId,
			songBpm.bpmMin,
			songBpm.bpmMax
		FROM songBpmVersion
		JOIN songBpm ON songBpmVersion.songBpmId = songBpm.songBpmId
	''')
	for songId, versionId, bpmMin, bpmMax in c.fetchall():
		songs[songId].bpm.add(versionId, Bpm(bpmMin, bpmMax))

	### Get song graphics
	c.execute('''
		SELECT
			songCard.songId,
			songCard.path,
			songCardVersion.versionId,
			operation.internalTitle
		FROM songCardVersion
		JOIN songCard ON songCardVersion.songCardId = songCard.songCardId
		JOIN operation ON songCardVersion.operationId = operation.operationId
	''')
	for songId, path, versionId, operation in c.fetchall():
		songs[songId].card.add(versionId, operation, path)

	return charts, songs, versions

def make_entry(charts, songs, versions, ver, chart):
	e = Entry()
	e.cid = chart.chartId
	e.sid = chart.songId
	#print("**** FOR DEBUGGING: CID=%d, SID=%d ****" % (e.cid,e.sid))
	e.stepmaker = chart.stepmaker
	e.cut = songs[e.sid].cut

	# This needs to be refactored
	vlist = sorted([v for v in versions.values()], key=lambda v: -v.sortOrder)
	for v in vlist:
		if chart.operations.recent(v.versionId, "DELETE") != "DELETE":
			e.last_seen = v
			break
	for v in vlist[::-1]:
		if chart.operations.recent(v.versionId, "DELETE") != "DELETE":
			e.first_seen = v
			break

	if ver == None:
		ver = e.last_seen.versionId

	e.rating = chart.rating.at(ver, Rating())
	e.comment = songs[e.sid].comment.at(ver, ("", None))[1]
	e.gameIdentifier = songs[e.sid].gameIdentifier.one(ver, "NOID")
	e.category = songs[e.sid].category.at(ver, "NOCATEGORY")
	e.bpm = songs[e.sid].bpm.at(ver, Bpm())
	e.card = songs[e.sid].card.best(ver, "NOCARD")
	e.labels = chart.labels.all(ver)
	e.title = songs[e.sid].title.at(ver, "NONAME")
	if e.title == "NONAME":
		e.title = songs[e.sid].fallbackTitle

	return e

def flatten_version(charts, songs, versions, ver, down=True):
	rows = []
	for chart in charts.values():
		if chart.operations.recent(ver, "DELETE") != "DELETE":
			e = make_entry(charts, songs, versions, ver, chart)
			rows.append(e)

	rows.sort(key=lambda e: (e.rating.order(down), e.cut.order(), e.title.lower()))
	return rows

def flatten_all(charts, songs, versions, down=True):
	rows = []
	for chart in charts.values():
		e = make_entry(charts, songs, versions, None, chart)
		rows.append(e)

	rows.sort(key=lambda e: (e.rating.order(down), e.cut.order(), e.title.lower()))
	return rows

def write_data_sheet(ws, rows):
	headers = [
		"CID",
		"SID",
		"GID",
		"Title",
		"Mode",
		"Difficulty",
		"Cut",
		"First Seen",
		"Last Seen",
		"BPM",
		"Category",
		"Stepmaker",
		"Labels",
		"Card",
		"Comment"
	]
	for i in range(len(headers)):
		ws.cell(row=1, column=i+1, value=headers[i]).font = Font(bold=True)
	for i in range(len(rows)):
		first_seen = last_seen = "?"
		if rows[i].first_seen != None:
			first_seen = str(rows[i].first_seen)
		if rows[i].last_seen != None:
			last_seen = str(rows[i].last_seen)
		ws.cell(row=i+2, column=1, value=rows[i].cid)
		ws.cell(row=i+2, column=2, value=rows[i].sid)
		ws.cell(row=i+2, column=3, value=rows[i].gameIdentifier)
		ws.cell(row=i+2, column=4, value=rows[i].title)
		ws.cell(row=i+2, column=5, value=rows[i].rating.mode_full())
		ws.cell(row=i+2, column=6, value=rows[i].rating.difficulty_str())
		ws.cell(row=i+2, column=7, value=str(rows[i].cut))
		ws.cell(row=i+2, column=8, value=first_seen)
		ws.cell(row=i+2, column=9, value=last_seen)
		ws.cell(row=i+2, column=10, value=str(rows[i].bpm))
		ws.cell(row=i+2, column=11, value=rows[i].category)
		ws.cell(row=i+2, column=12, value=str(rows[i].stepmaker))
		ws.cell(row=i+2, column=13, value=",".join(rows[i].labels))
		ws.cell(row=i+2, column=14, value=rows[i].card)
		ws.cell(row=i+2, column=15, value=rows[i].comment)

	ws.column_dimensions['A'].width = 5
	ws.column_dimensions['B'].width = 5
	ws.column_dimensions['C'].width = 5
	ws.column_dimensions['D'].width = 36
	ws.column_dimensions['E'].width = 9
	ws.column_dimensions['F'].width = 3
	ws.column_dimensions['G'].width = 9
	ws.column_dimensions['H'].width = 17
	ws.column_dimensions['I'].width = 17
	ws.column_dimensions['J'].width = 8
	ws.column_dimensions['K'].width = 14
	ws.column_dimensions['L'].width = 31
	ws.column_dimensions['M'].width = 39
	ws.column_dimensions['N'].width = 48
	ws.column_dimensions['O'].width = 120

	ws.freeze_panes = 'A2'

def write_score_sheet(ws, rows):
	headers = [
		"CID",           # A
		"Title",         # B
		"Mode",          # C
		"Difficulty",    # D
		"Cut",           # E
		"Passed (Pad)",  # F
		"Grade (Pad)",   # G
		"Miss (Pad)",    # H
		"Comment (Pad)", # I
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
		ws.cell(row=i+2, column=2, value="=VLOOKUP(A%d, Data!A1:O9999, 4, FALSE)" % (i+2)).fill = gray
		ws.cell(row=i+2, column=3, value="=VLOOKUP(A%d, Data!A1:O9999, 5, FALSE)" % (i+2)).fill = gray
		ws.cell(row=i+2, column=4, value="=VLOOKUP(A%d, Data!A1:O9999, 6, FALSE)" % (i+2)).fill = gray
		ws.cell(row=i+2, column=5, value="=VLOOKUP(A%d, Data!A1:O9999, 7, FALSE)" % (i+2)).fill = gray
		ws.cell(row=i+2, column=5).border = right
		ws.cell(row=i+2, column=9).border = right

	for i in range(len(rows)):
		ws.cell(row=i+2, column=9).alignment = Alignment(horizontal='fill')

	ws.column_dimensions['A'].width = 5
	ws.column_dimensions['B'].width = 30
	ws.column_dimensions['C'].width = 9
	ws.column_dimensions['D'].width = 3
	ws.column_dimensions['E'].width = 9
	ws.column_dimensions['F'].width = 4
	ws.column_dimensions['G'].width = 4
	ws.column_dimensions['H'].width = 4
	ws.column_dimensions['I'].width = 16
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

	ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"SSS"'], fill=grade_sss))
	ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"SS"'], fill=grade_ss))
	ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"S"'], fill=grade_s))
	ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"A"'], fill=grade_a))
	ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"B"'], fill=grade_b))
	ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"C"'], fill=grade_c))
	ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"D"'], fill=grade_d))
	ws.conditional_formatting.add('K2:K9999', CellIsRule(operator='equal', formula=['"F"'], fill=grade_f))

	ws.freeze_panes = 'C2'

def write_rows_dump(xlpath, rows, title):
	wb = Workbook()
	ws = wb.active
	ws.title = title
	write_data_sheet(ws, rows)
	wb.save(xlpath)

def write_rows_score(xlpath, rows, dbversion):
	wb = Workbook()
	ws_scores = wb.active
	ws_scores.title = "Scores"
	write_score_sheet(ws_scores, rows)

	ws_dump = wb.create_sheet(title="Data")
	write_data_sheet(ws_dump, rows)

	bold = Font(bold=True)
	ws_marker = wb.create_sheet(title="About")
	ws_marker.cell(row=1, column=1, value="Database Version:").font = bold
	ws_marker.cell(row=1, column=2, value=dbversion)
	ws_marker.cell(row=2, column=1, value="Sheet Generated On:").font = bold
	ws_marker.cell(row=2, column=2, value=str(datetime.datetime.now()))
	ws_marker.cell(row=3, column=1, value="Generator Version:").font = bold
	ws_marker.cell(row=3, column=2, value="v0.1")
	ws_marker.cell(row=5, column=1, value="Player:").font = bold
	ws_marker.cell(row=5, column=2, value="[YOUR NAME HERE]")
	ws_marker.column_dimensions['A'].width = 20
	ws_marker.column_dimensions['B'].width = 36

	wb.save(xlpath)

def dump_current(dbpath, xlpath):
	charts, songs, versions = read_database(dbpath)
	rows = flatten_version(charts, songs, versions, latest_version(versions).versionId)
	write_rows_dump(xlpath, rows, "Data for %s" % latest_version(versions))

def dump_all(dbpath, xlpath):
	charts, songs, versions = read_database(dbpath)
	rows = flatten_all(charts, songs, versions)
	write_rows_dump(xlpath, rows, "All data as of %s" % latest_version(versions))

def score_all(dbpath, xlpath):
	charts, songs, versions = read_database(dbpath)
	rows = flatten_all(charts, songs, versions)
	write_rows_score(xlpath, rows, dbpath)

def print_current(dbpath):
	charts, songs, versions = read_database(dbpath)
	rows = flatten_version(charts, songs, versions, latest_version(versions).versionId)
	for e in rows:
		print(str(e).encode())

def print_all(dbpath):
	charts, songs, versions = read_database(dbpath)
	rows = flatten_all(charts, songs, versions)
	for e in rows:
		print(str(e).encode())

if __name__ == '__main__':
	#print_current(DBPATH)
	#print_all(DBPATH)
	#dump_current(DBPATH, "output.xlsx")
	#dump_all(DBPATH, "output.xlsx")
	score_all(DBPATH, "output.xlsx")
