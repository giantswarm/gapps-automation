clasp_files = $(wildcard */.clasp.json)
gas_projects = $(subst /.clasp.json,/,$(clasp_files))
gas_projects_clean = $(subst /.clasp.json,-clean,$(clasp_files))

# unfortunately order matters here
# see: https://stackoverflow.com/questions/68379711/google-apps-script-hoisting-and-referenceerror
# (in Apps Script Editor one could set file position manually)
lib_files = lib/OAuth2.gs lib/UrlFetchJsonClient.js lib/CalendarListClient.js lib/PersonioAuthV1.js \
	lib/PersonioClientV1.js lib/GmailClientV1.js lib/SheetUtil.js lib/TriggerUtil.js lib/Util.js

.PHONY: all
all: $(gas_projects)
	@echo Assembled and pushed projects $^

%/: %/lib.js FORCE
	@echo Pushing project $@
	cd $@ && clasp push -f

.PRECIOUS: %/lib.js
%/lib.js: $(lib_files)
	@echo Updating library $@
	cat $^ > $@

.PHONY: clean
clean: $(gas_projects_clean)
	@echo Cleaned projects $(gas_projects)

# this target only exists to allow us to force pattern targets (.PHONY doesn't work there)
FORCE:
