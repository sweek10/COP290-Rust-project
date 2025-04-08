# Variables
CARGO := cargo
TARGET := target
BIN := spreadsheet

# Default target
all: build

# Build the project
build:
	$(CARGO) build

# Run the project
run: build
	$(CARGO) run

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
release:
	$(CARGO) build --release

# Remove the target directory
distclean: clean
	rm -rf $(TARGET)

.PHONY: all build run test clean fmt lint release distclean