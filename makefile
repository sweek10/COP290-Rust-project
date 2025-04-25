# Variables
CARGO := cargo
TARGET := target
BIN := spreadsheet
PORT := 80

# Default target
all: build

# Build the project (only build, generate binary named spreadsheet)
build:
	$(CARGO) build --release
	@if [ -f $(TARGET)/release/$(BIN) ]; then \
		echo "Binary generated at $(TARGET)/release/$(BIN)"; \
	else \
		echo "Error: Binary $(BIN) not found"; exit 1; \
	fi

# Run the project
run: build
	$(CARGO) run --release

# Test the project
test:
	$(CARGO) test

# Clean the build artifacts
clean:
	$(CARGO) clean

# Format the code
fmt:
	$(CARGO) fmt

# Check for linting issues
lint:
	$(CARGO) clippy -- -D warnings

# Release build
release: build

# Remove the target directory
distclean: clean
	rm -rf $(TARGET)

# Generate test coverage report
coverage:
	@if ! command -v cargo-tarpaulin >/dev/null 2>&1; then \
		echo "Installing cargo-tarpaulin..."; \
		$(CARGO) install cargo-tarpaulin; \
	fi
	cargo tarpaulin --out Html --output-dir $(TARGET)/coverage --timeout 120
	@echo "Coverage report generated at $(TARGET)/coverage/tarpaulin-report.html"

# Generate documentation (Rustdoc and PDF)
docs:
	$(CARGO) doc --no-deps
	@if ! command -v pandoc >/dev/null 2>&1; then \
		echo "pandoc is required for PDF generation. Please install it."; exit 1; \
	fi
	@if ! command -v wkhtmltopdf >/dev/null 2>&1; then \
		echo "wkhtmltopdf is required for PDF generation. Please install it."; exit 1; \
	fi
	@mkdir -p $(TARGET)/doc/pdf
	@echo "<html><body><h1>Spreadsheet Documentation</h1><iframe src='../index.html' width='100%' height='800px'></iframe></body></html>" > $(TARGET)/doc/index.html
	wkhtmltopdf $(TARGET)/doc/index.html $(TARGET)/doc/pdf/spreadsheet_docs.pdf
	@echo "Documentation generated at $(TARGET)/doc and PDF at $(TARGET)/doc/pdf/spreadsheet_docs.pdf"

# Run extension (Web Server/Client on localhost:80)
ext1: 
	@echo "Starting server on http://localhost:$(PORT)"
	$(CARGO) run --release -- --extension 999 18278 

.PHONY: all build run test clean fmt lint release distclean coverage docs ext1