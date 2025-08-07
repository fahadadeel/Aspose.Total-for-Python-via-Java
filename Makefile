.PHONY: help setup install test clean examples check-requirements

help:
	@echo "Clean Aspose Document Processor"
	@echo "==============================="
	@echo ""
	@echo "Available commands:"
	@echo "  setup           - Complete setup (recommended)"
	@echo "  install         - Install Python dependencies"
	@echo "  check-requirements - Check system requirements"
	@echo "  test           - Test the installation"
	@echo "  examples       - Run usage examples"
	@echo "  clean          - Clean generated files"

check-requirements:
	@echo "Checking system requirements..."
	@python3 --version || (echo "❌ Python 3 not found" && exit 1)
	@java -version || (echo "❌ Java not found. Install with: brew install openjdk@11" && exit 1)
	@echo "✅ System requirements OK"

install:
	@echo "Installing Python dependencies..."
	@pip install -r requirements.txt
	@echo "✅ Dependencies installed"

check-jars:
	@echo "Checking for Aspose JAR files..."
	@if ls plugins/*.jar 1> /dev/null 2>&1; then \
		echo "✅ Found JAR files:"; \
		ls -la plugins/*.jar; \
	else \
		echo "⚠️  No JAR files found"; \
		echo "📥 Download Aspose.Cells from: https://downloads.aspose.com/cells/java"; \
		echo "   Place the JAR file in plugins/ directory"; \
	fi

setup: check-requirements install check-jars
	@echo ""
	@echo "🎉 Setup complete!"
	@echo ""
	@echo "Next steps:"
	@echo "1. Download Aspose.Cells JAR if not already done"
	@echo "2. Run 'make test' to verify installation"
	@echo "3. Run 'make examples' to try conversions"
	@echo "4. Check 'examples/clean_usage.py' for API examples"

test:
	@echo "Testing installation..."
	@python3 -c "import sys; print('✅ Python import works')"
	@python3 -c "from main import AsposeDocumentProcessor; print('✅ Main module imports')"
	@python3 -c "from main import AsposeDocumentProcessor; p=AsposeDocumentProcessor(); print('✅ Processor initialized')"
	@echo "✅ Basic installation test passed"

examples:
	@echo "Running examples..."
	@cd examples && python3 clean_usage.py

clean:
	@echo "Cleaning generated files..."
	@find . -name "*.pyc" -delete
	@find . -name "__pycache__" -delete -type d 2>/dev/null || true
	@rm -f *.csv *.md *.json *.html
	@rm -f basic_output.* enhanced_output.* output.*
	@rm -rf batch_output/ test_batch/
	@echo "✅ Cleaned up"