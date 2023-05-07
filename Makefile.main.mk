clasp_files = $(wildcard */.clasp.json)
gas_projects = $(subst /.clasp.json,/,$(clasp_files))
gas_projects_clean = $(subst /.clasp.json,-clean,$(clasp_files))

# unfortunately order matters here
# see: https://stackoverflow.com/questions/68379711/google-apps-script-hoisting-and-referenceerror
# (in Apps Script Editor one could set file position manually)
#
# Header/trailer are set by target "lib" on demand.
# This is currently still needed for building the standalone lib (for other runtimes than GAS on v8)
# The variable lib_filter can be used to strip out certain patterns from library sources (like async/await).
#
lib_header_file =
lib_trailer_file =
lib_filter = 's/([[()\?\!\&\|,.;= +\-\*\t\t\n])(async|await)([() \t\n])/\1\3/g'
lib_files = lib/OAuth2.gs lib/UrlFetchJsonClient.js lib/CalendarListClient.js \
    lib/CalendarClient.js lib/PersonioAuthV1.js lib/PersonioClientV1.js lib/GmailClientV1.js lib/SheetUtil.js \
    lib/TriggerUtil.js lib/Util.js lib/PeopleTime.js

.PHONY: all
all: $(gas_projects)    ## Assemble and push all project using existing .clasp.json files (assumes all projects have been pushed before)
	@echo Assembled and pushed projects $^

.PHONY: lib
lib: ## Free standing library for use with NodeJS in lib-output/lib.js
lib: lib_header_file = lib/Header.js
lib: lib_trailer_file = lib/Trailer.js
lib: lib_filter = ''
lib: lib-output/lib.js

%/: %/lib.js FORCE  ## Assemble and push a single sub project (target name is sub directory name with trailing '/')
	@echo Pushing project $@
	@ if test -n "$(SCRIPT_ID)"; then \
   		 echo '{}' > $@.clasp.json ; \
 		 clasp-env --folder $@. --scriptId $(SCRIPT_ID); \
 	  else \
 	     echo WARNING: SCRIPT_ID not set, trying to use existing $@.clasp.json; \
 	  fi
	cd $@ && clasp push -f

#.PRECIOUS: %/lib.js
%/lib.js: $(lib_files)
	@echo Updating library $@
	mkdir $$(dirname $@) || true
	cat $(lib_header_file) $^ $(lib_trailer_file) | sed -r -E $(lib_filter) > $@

.PHONY: test
test: lib  ## Run tests using local NodeJS (node v16+ binary must be present)
	set -e ; cd tests && for t in *.*js; do node --input-type=module < $$t ; done

.PHONY: clean
clean:   ## Clean project
	rm -rf ./lib-output

# this target only exists to allow us to force pattern targets (.PHONY doesn't work there)
FORCE:
