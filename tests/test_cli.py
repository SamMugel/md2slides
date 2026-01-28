"""Tests for the CLI."""

import os
import tempfile

import pytest

from md2slides.cli import main


class TestCLI:
    """Test CLI functionality."""

    def test_main_with_valid_input(self):
        """CLI should succeed with valid input."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, "input.md")
            with open(input_path, "w") as f:
                f.write("# Test\n\n## Content\n\nHello")

            output_path = os.path.join(tmpdir, "output.pptx")
            result = main([input_path, "-o", output_path])

            assert result == 0
            assert os.path.exists(output_path)

    def test_main_with_default_output(self):
        """CLI should use default output path."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, "presentation.md")
            with open(input_path, "w") as f:
                f.write("# Test\n\n## Slide\n\nContent")

            result = main([input_path])

            assert result == 0
            assert os.path.exists(os.path.join(tmpdir, "presentation.pptx"))

    def test_main_with_missing_file(self):
        """CLI should return error for missing file."""
        result = main(["/nonexistent/file.md"])
        assert result == 1

    def test_main_with_invalid_markdown(self):
        """CLI should return error for invalid markdown."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, "invalid.md")
            with open(input_path, "w") as f:
                f.write("No headings here")

            result = main([input_path])
            assert result == 1

    def test_main_version(self, capsys):
        """CLI should show version."""
        with pytest.raises(SystemExit) as exc_info:
            main(["--version"])

        assert exc_info.value.code == 0
        captured = capsys.readouterr()
        assert "0.1.0" in captured.out
