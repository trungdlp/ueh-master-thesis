# UEH MBA Thesis build pipeline
# Usage:
#   make            -> build luanvan.docx
#   make open       -> build + open in Word
#   make reference  -> rebuild reference-ueh.docx style template
#   make figures    -> regenerate sample figures
#   make clean      -> remove generated artifacts

OUT_DIR  := outputs
DOCX     := $(OUT_DIR)/luanvan.docx
REF      := $(OUT_DIR)/reference-ueh.docx
BIB      := references.bib
CSL      := apa.csl
META     := metadata.yaml
SOURCES  := $(sort $(wildcard chapters/*.md))
PY       := ./venv/bin/python
PANDOC   := pandoc

.PHONY: all open reference figures clean

all: $(DOCX)

$(DOCX): $(SOURCES) $(REF) $(BIB) $(CSL) $(META) post_process.py | $(OUT_DIR)
	$(PANDOC) $(META) $(SOURCES) -o $(DOCX) \
		--reference-doc=$(REF) \
		--citeproc --csl=$(CSL) \
		--from markdown+raw_attribute
	$(PY) post_process.py $(DOCX)

$(REF): build_reference.py | $(OUT_DIR)
	$(PY) build_reference.py $(REF)

$(OUT_DIR):
	mkdir -p $(OUT_DIR)

reference: $(REF)

figures:
	$(PY) make_figures.py

open: $(DOCX)
	open $(DOCX)

clean:
	rm -rf $(OUT_DIR)
