clasp_files = $(wildcard */.clasp.json)
lib_files = $(wildcard lib/*.js)
gas_projects = $(subst /.clasp.json,/,$(clasp_files))
gas_projects_clean = $(subst /.clasp.json,-clean,$(clasp_files))

.PHONY: all
all: $(gas_projects)
	@echo Assembled and pushed projects $^

%/: %/lib/ FORCE
	@echo Pushing project $@
	cd $@ && clasp push -f

%/lib/: $(lib_files)
	@echo Updating library $@
	rm -rf "$@" ; cp -a ./lib/. "$@"


.PHONY: clean
clean: $(gas_projects_clean)
	@echo Cleaned projects $(gas_projects)

.PHONY: %-clean
%-clean:
	@echo Cleaning $(subst -clean,/,$@)
	rm -rf "$(subst -clean,/lib,$@)"


# this target only exists to allow us to force pattern targets (.PHONY doesn't work there)
FORCE:
