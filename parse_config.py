SECTION_NONE = 0
SECTION_MIXES = 1
SECTION_MODES = 2
SECTION_DIFFICULTIES = 3
SECTION_MISC = 4

class ConfigError(Exception):
	pass

class Config:
	def __init__(self):
		self.mixes = []
		self.modes = []
		self.diff_min = 1
		self.diff_max = 30
		self.unrated = True
		self.pad = True
		self.keyboard = True

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

	return config

def verify_config(config, mixes, modes):
	for m in config.mixes:
		if not m in mixes:
			raise ConfigError("mix %s not in database" % m)

	for m in config.modes:
		if not m in modes:
			raise ConfigError("mode %s not in database %s" % (m, repr(modes)))

	if config.diff_min > config.diff_max:
		raise ConfigError("minimum difficulty (%d) > maximum difficulty (%d)" % (config.diff_min, config.diff_max))
