.PHONY: all install lint test audit
all: lint test
install:
	python -m pip install -r requirements.txt
lint:
	python -m ruff check tests
test:
	python -m pytest -q tests
audit:
	python belajar/tools/verify_repository.py
