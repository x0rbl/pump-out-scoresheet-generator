import sqlite3

LANGUAGE = "en"

OP_NONE   = 0
OP_INSERT = 1
OP_DELETE = 2
OP_UPDATE = 3
OP_EXISTS = 4
OP_REVIVE = 5
OP_CROSS  = 6

class ParseError(Exception):
	pass

class Mix:
	def __init__(self, mixId, mixTitle, parentId, sortOrder):
		self.id = mixId
		self.title = mixTitle
		self.parent = parentId
		self.order = sortOrder

class Version:
	def __init__(self, versionId, mixId, versionTitle, parentId, sortOrder):
		self.id = versionId
		self.mix = mixId
		self.title = versionTitle
		self.parent = parentId
		self.order = sortOrder

class VersionedValue:
	def __init__(self):
		# Maps versionId -> value
		self.values = {}
		# A list of the versionIds in self.values, sorted BACKWARDS by order.
		# Or None if the cache needs to be rebuilt (use ensure_cache).
		self.cache = None

	def add(self, versionId, value):
		if versionId in self.values:
			raise Exception("duplicate key: vid=%d, new=%s, old=%s" % (versionId, value, self.values[versionId]))
		self.values[versionId] = value
		self.cache = None

	def ensure_cache(self, versions):
		if self.cache == None:
			self.cache = sorted([vid for vid in self.values], key=lambda e: -versions[e].order)

	def get_list(self, versions):
		self.ensure_cache(versions)
		return [(v, self.values[v]) for v in self.cache]

class MultipleVersionedValue:
	def __init__(self):
		self.values = {}

	def add(self, versionId, operation, value):
		if not value in self.values:
			self.values[value] = VersionedValue()
		self.values[value].add(versionId, operation)

class Mode:
	def __init__(self, modeId, modeTitle, abbreviation, color, sortOrder, padsUsed, routine, coOp, performance):
		self.id = modeId
		self.title = modeTitle
		self.abbr = abbreviation
		self.color = color
		self.order = sortOrder
		self.pads = padsUsed
		self.routine = routine
		self.coop = coOp
		self.performance = performance

class Cut:
	def __init__(self, cutId, cutTitle, sortOrder):
		self.id = cutId
		self.title = cutTitle
		self.order = sortOrder

class Rating:
	def __init__(self, modeId=None, difficulty=None):
		self.mode = modeId
		self.difficulty = difficulty

class Bpm:
	def __init__(self, low=None, high=None):
		self.low = low
		self.high = high

	def __str__(self):
		if self.low == self.high == None:
			return "NOBPM"
		if self.low == self.high:
			if int(self.low) == self.low:
				return "%d" % int(self.low)
			return "%s" % self.low
		if int(self.low) == self.low and int(self.high) == self.high:
			return "%d-%d" % (self.low, self.high)
		return "%s-%s" % (self.low, self.high)

class NameGroup:
	def __init__(self):
		self.str = ""
		self.list = []

	def add(self, prefix, name):
		if self.str != "":
			if prefix != "":
				self.str += " " + prefix
			else:
				self.str += ","

		if self.str != "" and name != "":
			self.str += " "
		self.str += name

		self.list.append(name)

	def __str__(self):
		return self.str

class Song:
	def __init__(self):
		self.songId = -1
		self.operations = VersionedValue()
		self.title = VersionedValue()
		self.gameIdentifier = MultipleVersionedValue()
		self.category = VersionedValue()
		self.bpm = VersionedValue()
		self.card = MultipleVersionedValue()
		self.artist = NameGroup()
		self.cut = -1
		self.fallbackTitle = "NOTITLE"

class Chart:
	def __init__(self):
		self.chartId = -1
		self.songId = -1
		self.operations = VersionedValue()
		self.rating = VersionedValue()
		self.stepmaker = NameGroup()
		self.labels = MultipleVersionedValue()

class Database:
	def __init__(self):
		self.mixes = {}
		self.modes = {}
		self.cuts = {}
		self.versions = {}
		self.charts = {}
		self.songs = {}
		self.ratingImages = {}

	# Returns the value in a VersionedValue that existed at versionId
	# Returns default if no value is set at versionId
	#
	# Searching works chronologically without regard for whether the most recent update
	# occurred in a separate branch of the version tree
	def _vv_at(self, versionedValue, versionId, default):
		version, value = self._vv_at_version(versionedValue, versionId, default)
		return value

	# Returns the value in a VersionedValue that existed at versionId, along with
	# the version id of the version that set that value: (versionId, value)
	# Returns (None, default) if no value is set at versionId
	#
	# Searching works chronologically without regard for whether the most recent update
	# occurred in a separate branch of the version tree
	def _vv_at_version(self, versionedValue, versionId, default):
		if not versionId in self.versions:
			return default
		so = self.versions[versionId].order
		versionedValue.ensure_cache(self.versions)
		for ver in versionedValue.cache:
			if self.versions[ver].order <= so:
				return (ver, versionedValue.values[ver])
		return (None, default)

	# Returns the value in a VersionedValue, from within the same version tree, that existed
	# at versionId
	# Returns default if no value is set at versionId
	def _vv_recent(self, versionedValue, versionId, default):
		while True:
			if not versionId in self.versions:
				return default
			if not versionId in versionedValue.values:
				versionId = self.versions[versionId].parent
				continue
			return versionedValue.values[versionId]

	# Returns the (versionId,value) associated with the earliest chronological value
	# in a VersionedValue
	# If no elements exist, returns (None, default)
	def _vv_earliest(self, versionedValue, default):
		if len(versionedValue.values) == 0:
			return (None, default)
		versionedValue.ensure_cache(self.versions)
		ver = versionedValue.cache[-1]
		return (ver, versionedValue.values[ver])

	# Returns the (versionId,value) associated with the latest chronological value
	# in a VersionedValue
	# If no elements exist, returns (None, default)
	def _vv_latest(self, versionedValue, default):
		if len(versionedValue.values) == 0:
			return (None, default)
		versionedValue.ensure_cache(self.versions)
		ver = versionedValue.cache[0]
		return (ver, versionedValue.values[id])

	# Returns all values in a MultipleVersionedValue that exist at versionId
	def _mvv_all(self, multiVersionedValue, versionId):
		res = []
		for value, vv in multiVersionedValue.values.items():
			if self._vv_at(vv, versionId, OP_DELETE) != OP_DELETE:
				res.append(value)
		return res

	# Returns the one value that exists in a MultipleVersionedValue at versionId
	# If no value exists at versionId, return default
	# If more than one value exists at versionId, throw an error
	# If this function is throwing an error due to a possible problem in the data, use _mvv_best
	def _mvv_one(self, multiVersionedValue, versionId, default):
		vals = self._mvv_all(multiVersionedValue, versionId)
		if len(vals) > 1:
			raise Exception("too many vals: %s" % vals)
		if vals == []:
			return default
		return vals[0]

	# Similar to _mvv_one, except if multiple values exist at versionId, return the one with
	# the newest version
	def _mvv_best(self, multiVersionedValue, versionId, default):
		pairs = []
		for value, vv in multiVersionedValue.values.items():
			(version_at, op_at) = self._vv_at_version(vv, versionId, OP_DELETE)
			if op_at != OP_DELETE:
				pairs.append((version_at, value))

		if len(pairs) == 0:
			return default

		best_so = -1
		best_val = None
		for (ver,val) in pairs:
			so = self.versions[ver].order
			if so > best_so:
				best_so = so
				best_val = val
		return val

	def latest_version(self):
		best_id = None
		best_so = -1
		for ver in self.versions.values():
			if ver.order > best_so:
				best_so = ver.order
				best_id = ver.id
		return best_id

	def newest_version_from_mix(self, mixId):
		vlist = sorted([v for v in self.versions.values()], key=lambda e: -e.order)
		for ver in vlist:
			if ver.mix == mixId:
				return ver.id
		return None

	def chart_in_mix(self, chartId, mixId):
		return self.chart_version_in_mix(chartId, mixId) != None

	def chart_version_in_mix(self, chartId, mixId):
		if not chartId in self.charts:
			return None
		vlist = sorted([v for v in self.versions.values()], key=lambda e: -e.order)
		for ver in vlist:
			if ver.mix == mixId:
				(op, comment) = self._vv_recent(self.charts[chartId].operations, ver.id, (OP_DELETE, None))
				if op != OP_DELETE:
					return ver.id
		return None

	def chart_sort_key(self, chartId, versionId, down=True):
		rating = self.chart_rating(chartId, versionId)
		title = self.song_title(self.chart_song(chartId), versionId).lower()

		mode_key = 99
		if rating != None and rating.mode != None:
			mt = self.modes[rating.mode].title.lower()
			if "single" in mt:
				mode_key = 0
			elif "half" in mt:
				mode_key = 2
			elif "double" in mt:
				mode_key = 1
			elif "routine" in mt:
				mode_key = 3
			elif "co" in mt:
				mode_key = 4

		diff_key = 99
		if rating.difficulty != None:
			diff_key = rating.difficulty
			if down:
				diff_key = -diff_key

		return (diff_key, mode_key, title)

	def chart_rating(self, chartId, versionId):
		if not chartId in self.charts:
			return None
		return self._vv_at(self.charts[chartId].rating, versionId, None)

	def chart_rating_sequence_str(self, chartId, changes_only=False):
		ratings = self.charts[chartId].rating.get_list(self.versions)[::-1]

		elems = []
		for i in range(len(ratings)):
			(v,r) = ratings[i]
			rs = self.rating_str(r)
			if len(elems) == 0 or elems[-1] != rs:
				elems.append(rs)

		if changes_only and len(elems) <= 1:
			return ""

		return " -> ".join(elems)

	def chart_mode(self, chartId, versionId):
		rating = self.chart_rating(chartId, versionId)
		if rating == None:
			return None
		return rating.mode

	def chart_mode_str(self, chartId, versionId):
		mode = self.chart_mode(chartId, versionId)
		if mode == None:
			return None
		return self.modes[mode].title

	def chart_difficulty(self, chartId, versionId):
		rating = self.chart_rating(chartId, versionId)
		if rating == None:
			return None
		return rating.difficulty

	def chart_difficulty_str(self, chartId, versionId):
		difficulty = self.chart_difficulty(chartId, versionId)
		if difficulty == None: return "??"
		return "%02d" % difficulty

	def chart_introduced(self, chartId):
		if not chartId in self.charts:
			return None
		(ver, (op, comment)) = self._vv_earliest(self.charts[chartId].operations, (OP_NONE, None))
		if op == OP_INSERT:
			return ver
		return None

	def chart_last_seen(self, chartId):
		if not chartId in self.charts:
			return None
		vlist = sorted([v for v in self.versions.values()], key=lambda v: -v.order)
		for ver in vlist:
			(op, comment) = self._vv_recent(self.charts[chartId].operations, ver.id, (OP_DELETE, None))
			if op != OP_DELETE:
				return ver.id
		return None

	def chart_song(self, chartId):
		if not chartId in self.charts:
			return None
		return self.charts[chartId].songId

	def song_game_id(self, songId, versionId):
		if not songId in self.songs:
			return None
		return self._mvv_one(self.songs[songId].gameIdentifier, versionId, None)

	def chart_stepmaker(self, chartId):
		if not chartId in self.charts:
			return None
		return self.charts[chartId].stepmaker

	def chart_labels(self, chartId, versionId):
		if not chartId in self.charts:
			return []
		return self._mvv_all(self.charts[chartId].labels, versionId)

	def chart_comment(self, chartId, versionId):
		if not chartId in self.charts:
			return None
		(op, comment) = self._vv_at(self.charts[chartId].comment, versionId, (OP_NONE, None))
		return comment

	def song_title(self, songId, versionId):
		if not songId in self.songs:
			return None
		title = self._vv_at(self.songs[songId].title, versionId, None)
		if title != None:
			return title
		return self.songs[songId].fallbackTitle

	def song_cut(self, songId):
		if not songId in self.songs:
			return None
		cutId = self.songs[songId].cut
		return self.cuts.get(cutId, None)

	def song_cut_str(self, songId):
		cut = self.song_cut(songId)
		if cut == None: return "NOCUT"
		return cut.title

	def song_bpm(self, songId, versionId):
		if not songId in self.songs:
			return None
		return self._vv_at(self.songs[songId].bpm, versionId, None)

	def song_bpm_str(self, songId, versionId):
		bpm = self.song_bpm(songId, versionId)
		if bpm == None: return "NOBPM"
		return str(bpm)

	def song_category(self, songId, versionId):
		if not songId in self.songs:
			return None
		return self._vv_at(self.songs[songId].category, versionId, "NOCATEGORY")

	def song_card(self, songId, versionId):
		if not songId in self.songs:
			return None
		# card data seems to be bugged, so use _mvv_best instead of _mvv_one
		return self._mvv_best(self.songs[songId].card, versionId, "NOCARD")

	def song_comment(self, songId, versionId):
		if not songId in self.songs:
			return None
		(op, comment) = self._vv_at(self.songs[songId].operations, versionId, (OP_NONE, None))
		return comment

	def version_title(self, versionId):
		if not versionId in self.versions:
			return None
		mixId = self.versions[versionId].mix
		return "%s %s" % (self.mixes[mixId].title, self.versions[versionId].title)

	def rating_str(self, rating):
		if rating == None:
			return "???"
		m = "?"
		if rating.mode != None:
			m = self.modes[rating.mode].abbr
		d = "??"
		if rating.difficulty != None:
			d = "%02d" % rating.difficulty
		return m+d

def read_database(dbpath):
	conn = sqlite3.connect(dbpath)
	c = conn.cursor()

	db = Database()

	c.execute('''
		SELECT
			operationId,
			internalTitle
		FROM operation
	''')
	results = c.fetchall()
	expect = [
		(OP_INSERT, "INSERT"),
		(OP_DELETE, "DELETE"),
		(OP_UPDATE, "UPDATE"),
		(OP_EXISTS, "EXISTS"),
		(OP_REVIVE, "REVIVE"),
		(OP_CROSS,  "CROSS"),
	]
	if results != expect:
		raise ParseError("incompatible operations list")

	# Get mixes
	c.execute('''
		SELECT
			mixId,
			internalTitle,
			parentMixId,
			sortOrder
		FROM mix
	''')
	for mixId, mixTitle, parentId, sortOrder in c.fetchall():
		db.mixes[mixId] = Mix(mixId, mixTitle, parentId, sortOrder)

	# Get modes
	c.execute('''
		SELECT
			modeId,
			internalTitle,
			internalAbbreviation,
			internalHexColor,
			sortOrder,
			padsUsed,
			routine,
			coOp,
			performance
		FROM mode
	''')
	for modeId, title, abbreviation, color, sortOrder, padsUsed, routine, coOp, performance in c.fetchall():
		db.modes[modeId] = Mode(modeId, title, abbreviation, color, sortOrder, padsUsed, routine, coOp, performance)

	# Get cuts
	c.execute('''
		SELECT
			cutId,
			internalTitle,
			sortOrder
		FROM cut
	''')
	for cutId, cutTitle, sortOrder in c.fetchall():
		db.cuts[cutId] = Cut(cutId, cutTitle, sortOrder)

	### Get versions
	c.execute('''
		SELECT
			versionId,
			mixId,
			internalTitle,
			parentVersionId,
			sortOrder
		FROM version
	''')
	for versionId, mixId, versionTitle, parentId, sortOrder in c.fetchall():
		db.versions[versionId] = Version(versionId, mixId, versionTitle, parentId, sortOrder)

	### Populate charts by version
	c.execute('''
		SELECT
			chartVersion.chartId,
			chart.songId,
			chartVersion.versionId,
			chartVersion.operationId,
			chartVersion.internalDescription
		FROM chartVersion
		JOIN chart ON chart.chartId = chartVersion.chartId
	''')
	for chartId, songId, versionId, operationId, comment in c.fetchall():
		if not chartId in db.charts:
			db.charts[chartId] = Chart()
			db.charts[chartId].chartId = chartId
			db.charts[chartId].songId = songId
		db.charts[chartId].operations.add(versionId, (operationId, comment))

	### Get chart ratings
	c.execute('''
		SELECT
			chartRatingVersion.chartId,
			chartRatingversion.versionId,
			chartRating.modeId,
			difficulty.value
		FROM chartRatingVersion
		JOIN chartRating on chartRatingVersion.chartRatingId = chartRating.chartRatingId
		JOIN difficulty ON chartRating.difficultyId = difficulty.difficultyId
	''')
	for chartId, versionId, mode, difficulty in c.fetchall():
		db.charts[chartId].rating.add(versionId, Rating(mode, difficulty))

	### Get rating paths
	c.execute('''
		SELECT
			rating.modeId,
			difficulty.value,
			rating.path
		FROM rating
		JOIN difficulty ON difficulty.difficultyId = rating.difficultyId
	''')
	for modeId, difficulty, path in c.fetchall():
		db.ratingImages[(modeId, difficulty)] = path

	### Get chart stepmakers
	c.execute('''
		SELECT
			chartStepmaker.chartId,
			chartStepmaker.prefix,
			stepmaker.internalTitle,
			chartStepmaker.sortOrder
		FROM chartStepmaker
		JOIN stepmaker ON chartStepmaker.stepmakerId = stepmaker.stepmakerId
		ORDER BY chartStepmaker.chartId ASC, chartStepmaker.sortOrder ASC
	''')
	for chartId, prefix, stepmaker, _ in c.fetchall():
		db.charts[chartId].stepmaker.add(prefix, stepmaker)

	### Get chart labels
	c.execute('''
		SELECT
			chartLabel.chartId,
			chartLabelVersion.versionId,
			chartLabelVersion.operationId,
			label.internalTitle
		FROM chartLabelVersion
		JOIN chartLabel ON chartLabelVersion.chartLabelId = chartLabel.chartLabelId
		JOIN label ON chartLabel.labelId = label.labelId
	''')
	for chartId, versionId, operationId, label in c.fetchall():
		db.charts[chartId].labels.add(versionId, operationId, label)

	### Create songs by version, with cut (Full Song, Remix, etc) and fallback title
	c.execute('''
		SELECT
			songVersion.songId,
			songVersion.versionId,
			songVersion.operationId,
			songVersion.internalDescription,
			song.cutId,
			song.internalTitle
		FROM songVersion
		JOIN song ON songVersion.songId = song.songId
	''')
	for songId, versionId, operationId, comment, cutId, fallbackTitle in c.fetchall():
		if not songId in db.songs:
			db.songs[songId] = Song()
			db.songs[songId].songId = songId
		db.songs[songId].operations.add(versionId, (operationId, comment))
		db.songs[songId].cut = cutId
		db.songs[songId].fallbackTitle = fallbackTitle

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
		WHERE language.code = "%s"
	''' % LANGUAGE)
	for songId, versionId, title in c.fetchall():
		db.songs[songId].title.add(versionId, title)

	### Get official song identifiers
	c.execute('''
		SELECT
			songGameIdentifier.songId,
			songGameIdentifier.gameIdentifier,
			songGameIdentifierVersion.versionId,
			songGameIdentifierVersion.operationId
		FROM songGameIdentifierVersion
		JOIN songGameIdentifier ON songGameIdentifierVersion.songGameIdentifierId = songGameIdentifier.songGameIdentifierId
	''')
	for songId, gameIdentifier, versionId, operationId in c.fetchall():
		db.songs[songId].gameIdentifier.add(versionId, operationId, gameIdentifier)

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
		db.songs[songId].category.add(versionId, category)

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
		db.songs[songId].bpm.add(versionId, Bpm(bpmMin, bpmMax))

	### Get song artists
	c.execute('''
		SELECT
			songArtist.songId,
			songArtist.prefix,
			artist.internalTitle,
			songArtist.sortOrder
		FROM songArtist
		JOIN artist ON songArtist.artistId = artist.artistId
		ORDER BY songArtist.songId ASC, songArtist.sortOrder ASC
	''')
	for songId, prefix, artist, _ in c.fetchall():
		db.songs[songId].artist.add(prefix, artist)

	### Get song graphics
	c.execute('''
		SELECT
			songCard.songId,
			songCard.path,
			songCardVersion.versionId,
			songCardVersion.operationId
		FROM songCardVersion
		JOIN songCard ON songCardVersion.songCardId = songCard.songCardId
	''')
	for songId, path, versionId, operationId in c.fetchall():
		db.songs[songId].card.add(versionId, operationId, path)

	return db
