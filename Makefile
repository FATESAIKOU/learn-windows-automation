# Windows Automation Toolkit Makefile

.PHONY: install install-dev test test-coverage lint clean help

# Default target
help:
	@echo "Available targets:"
	@echo "  install        - Install the package dependencies"
	@echo "  install-dev    - Install development dependencies"
	@echo "  test          - Run all tests"
	@echo "  test-coverage - Run tests with coverage report"
	@echo "  lint          - Run code linting (if configured)"
	@echo "  clean         - Clean build artifacts"
	@echo "  run-win32     - Run win32 automation demo"
	@echo "  run-pywinauto - Run pywinauto automation demo"

# Install dependencies
install:
	uv sync

# Install development dependencies
install-dev:
	uv sync --dev

# Run tests
test:
	uv run pytest

# Run tests with coverage
test-coverage:
	uv run pytest --cov=learn_windows_automation --cov-report=html --cov-report=term

# Clean build artifacts
clean:
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info/
	rm -rf htmlcov/
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete

# Run win32 automation
run-win32:
	uv run python main.py win32

# Run pywinauto automation
run-pywinauto:
	uv run python main.py pywinauto

# Build the package
build:
	uv build

# Install in editable mode for development
dev-install:
	uv pip install -e .
