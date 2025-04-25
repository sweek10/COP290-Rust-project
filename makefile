# Default: build the autograder binary
all: build

# Build autograder binary with the autograder feature
build:
	cargo build --release

# Run the autograder binary (100×100 grid)
run: build
	./target/release/spreadsheet 999 18278

# Build & run the web server extension (requires sudo for port 80)
ext1:
	cargo build --release
	./target/release/spreadsheet --extension 999 18278

# Clean everything
clean:
	cargo clean
	rm -rf dist build

coverage:
	cargo tarpaulin --out Html

test:
	cargo test

docs: cargo-doc report

cargo-doc:
	cargo doc --no-deps --open

report: report.pdf
	# On macOS, use 'open'; on Linux, you might use 'xdg-open'
	open report.pdf

report.pdf: report.tex
	pdflatex report.tex
	# Running pdflatex a second time for proper reference resolution (optional)
	pdflatex report.tex

.PHONY: all build run extension clean coverage test docs cargo-doc report
