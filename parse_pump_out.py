import sqlite3

class Mix:
	def __init__(self, mixId, mixTitle, parentId, sortOrder):
		self.id = mixId
		self.title = mixTitle
		self.parent = parentId
		self.sortOrder = sortOrder

	def __str__(self):
		return self.title

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
		# TODO: Make a mode object and pass that instead of initializing with
		# the abbreviation. This no longer makes sense now that we are
		# filtering on mode names in the config.
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
		if self.low == self.high == -1.0:
			return "NOBPM"
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

class Database:
	def __init__(self, versions, mixes, charts, songs, modes):
		self.versions = versions
		self.mixes = mixes
		self.charts = charts
		self.songs = songs
		self.modes = modes

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
		self.mixes = set()

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

	# Get mix tree
	c.execute('''
		SELECT
			mix.mixId,
			mix.internalTitle,
			mix.parentMixId,
			mix.sortOrder
		FROM mix
	''')
	mixes = {}
	for mixId, mixTitle, parentId, sortOrder in c.fetchall():
		mixes[mixId] = Mix(mixId, mixTitle, parentId, sortOrder)

	### Get version tree
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

	### Get list of mode names
	c.execute('''
		SELECT mode.internalTitle
		FROM mode
	''')
	modes = set()
	for mode in c.fetchall():
		modes.add(mode[0])

	return Database(versions, mixes, charts, songs, modes)

def make_entry(db, chart):
	e = Entry()
	e.cid = chart.chartId
	e.sid = chart.songId
	#print("**** FOR DEBUGGING: CID=%d, SID=%d ****" % (e.cid,e.sid))
	e.stepmaker = chart.stepmaker
	e.cut = db.songs[e.sid].cut

	# This needs to be refactored
	e.first_seen = None
	e.last_seen = None
	vlist = sorted([v for v in db.versions.values()], key=lambda v: v.sortOrder)
	for v in vlist:
		op = chart.operations.recent(v.versionId, "DELETE")
		if op != "DELETE":
			e.mixes.add(v.mixTitle)
			if op == "INSERT" and e.first_seen == None:
				e.first_seen = v
			e.last_seen = v

	ver = e.last_seen.versionId

	e.rating = chart.rating.at(ver, Rating())
	e.comment = db.songs[e.sid].comment.at(ver, ("", None))[1]
	e.gameIdentifier = db.songs[e.sid].gameIdentifier.one(ver, "NOID")
	e.category = db.songs[e.sid].category.at(ver, "NOCATEGORY")
	e.bpm = db.songs[e.sid].bpm.at(ver, Bpm())
	e.card = db.songs[e.sid].card.best(ver, "NOCARD")
	e.labels = chart.labels.all(ver)
	e.title = db.songs[e.sid].title.at(ver, "NONAME")
	if e.title == "NONAME":
		e.title = db.songs[e.sid].fallbackTitle

	return e

def flatten_db(db, down=True):
	rows = []
	for chart in db.charts.values():
		e = make_entry(db, chart)
		rows.append(e)

	rows.sort(key=lambda e: (e.rating.order(down), e.cut.order(), e.title.lower()))
	return rows