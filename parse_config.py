SECTION_NONE = 0
SECTION_MIXES = 1
SECTION_MODES = 2
SECTION_DIFFICULTIES = 3
SECTION_MISC = 4

class ConfigError(Exception):
	pass

class Config:
	def __init__(self):
		# These options are set by parse_config
		self.mixes = []
		self.modes = []
		self.diff_min = 1
		self.diff_max = 30
		self.unrated = True
		self.pad = True
		self.keyboard = True
		# These options need to be set manually via titles_to_ids
		self.mix_ids = []
		self.mode_ids = []

def parse_config(config_path):
	config = Config()

	section = SECTION_NONE
	for lineb in open(config_path, "rb").read().splitlines():
		line = lineb.decode()
		if '#' in line:
			line = line[:line.find('#')]
		line = line.strip()
		if line == "": continue

		if line == "[Mixes]":
			section = SECTION_MIXES
			continue
		if line == "[Modes]":
			section = SECTION_MODES
			continue
		if line == "[Difficulties]":
			section = SECTION_DIFFICULTIES
			continue
		if line == "[Misc]":
			section = SECTION_MISC
			continue

		if section == SECTION_MIXES:
			config.mixes.append(line)
		elif section == SECTION_MODES:
			config.modes.append(line)
		elif section == SECTION_DIFFICULTIES:
			if line.startswith("Min="):
				config.diff_min = int(line[len("Min="):])
			elif line.startswith("Max="):
				config.diff_max = int(line[len("Max="):])
			elif line.startswith("IncludeUnrated="):
				config.unrated = int(line[len("IncludeUnrated="):]) != 0
			else:
				raise ConfigError("unsupported option in [Difficulties] section: %s" % line)
		elif section == SECTION_MISC:
			if line.startswith("IncludePad="):
				config.pad = int(line[len("IncludePad="):]) != 0
			elif line.startswith("IncludeKbd="):
				config.keyboard = int(line[len("IncludeKbd="):]) != 0
			else:
				raise ConfigError("unsupported option in [Misc] section: %s" % line)

	if config.diff_min > config.diff_max:
		raise ConfigError("minimum difficulty (%d) > maximum difficulty (%d)" % (config.diff_min, config.diff_max))

	if len(config.mixes) == 0:
		raise ConfigError("no mixes specified")
	if len(config.modes) == 0:
		raise ConfigError("no mixes specified")

	return config

def titles_to_ids(titles, collection, what):
	title_to_id = {val.title: id for (id, val) in collection.items()}
	for t in titles:
		if not t in title_to_id:
			raise ConfigError("%s %s not in database" % (what, t))
	return list(map(lambda t: title_to_id[t], titles))

def config_all(db):
	config = Config()

	for mid in sorted(db.mixes, key=lambda e: db.mixes[e].order):
		config.mix_ids.append(mid)
		config.mixes.append(db.mixes[mid].title)

	for mid in sorted(db.modes, key=lambda e: db.modes[e].order):
		config.mode_ids.append(mid)
		config.modes.append(db.modes[mid].title)

	config.diff_min = -float("inf")
	config.diff_max = float("inf")
	config.unrated = True

	config.pad = True
	config.keyboard = True

	return config
