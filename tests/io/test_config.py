"""Tests for configuration module."""

import tempfile
from pathlib import Path
from unittest.mock import patch


from lotus123.ui.config import AppConfig


class TestAppConfigDefault:
    """Tests for AppConfig defaults."""

    def test_default_values(self):
        """Test default configuration values."""
        config = AppConfig()
        assert config.theme == "LOTUS"
        assert config.default_col_width == 10
        assert config.recent_files == []

    def test_custom_values(self):
        """Test custom configuration values."""
        config = AppConfig(theme="MOCHA", default_col_width=15, recent_files=["/path/to/file.wk1"])
        assert config.theme == "MOCHA"
        assert config.default_col_width == 15
        assert config.recent_files == ["/path/to/file.wk1"]


class TestAppConfigRecentFiles:
    """Tests for recent files functionality."""

    def test_add_recent_file(self):
        """Test adding a recent file."""
        config = AppConfig()
        config.add_recent_file("/path/to/file1.wk1")
        assert config.recent_files == ["/path/to/file1.wk1"]

    def test_add_multiple_recent_files(self):
        """Test adding multiple recent files."""
        config = AppConfig()
        config.add_recent_file("/path/to/file1.wk1")
        config.add_recent_file("/path/to/file2.wk1")
        config.add_recent_file("/path/to/file3.wk1")
        assert config.recent_files == [
            "/path/to/file3.wk1",
            "/path/to/file2.wk1",
            "/path/to/file1.wk1",
        ]

    def test_add_duplicate_recent_file(self):
        """Test adding a duplicate file moves it to front."""
        config = AppConfig()
        config.add_recent_file("/path/to/file1.wk1")
        config.add_recent_file("/path/to/file2.wk1")
        config.add_recent_file("/path/to/file1.wk1")
        assert config.recent_files == ["/path/to/file1.wk1", "/path/to/file2.wk1"]

    def test_recent_files_max_10(self):
        """Test recent files are limited to 10."""
        config = AppConfig()
        for i in range(15):
            config.add_recent_file(f"/path/to/file{i}.wk1")
        assert len(config.recent_files) == 10
        assert config.recent_files[0] == "/path/to/file14.wk1"


class TestAppConfigSaveLoad:
    """Tests for save/load functionality."""

    def test_save_and_load(self):
        """Test saving and loading configuration."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_dir = Path(tmpdir)
            config_file = config_dir / "config.json"

            # Mock CONFIG_DIR and CONFIG_FILE
            with patch("lotus123.ui.config.CONFIG_DIR", config_dir):
                with patch("lotus123.ui.config.CONFIG_FILE", config_file):
                    # Create and save config
                    config = AppConfig(
                        theme="MOCHA", default_col_width=12, recent_files=["/test/file.wk1"]
                    )
                    config.save()

                    # Verify file was created
                    assert config_file.exists()

                    # Load and verify
                    loaded = AppConfig.load()
                    assert loaded.theme == "MOCHA"
                    assert loaded.default_col_width == 12
                    assert loaded.recent_files == ["/test/file.wk1"]

    def test_load_missing_file(self):
        """Test loading when config file doesn't exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "nonexistent.json"

            with patch("lotus123.ui.config.CONFIG_FILE", config_file):
                loaded = AppConfig.load()
                # Should return default config
                assert loaded.theme == "LOTUS"
                assert loaded.default_col_width == 10

    def test_load_invalid_json(self):
        """Test loading with invalid JSON."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_file = Path(tmpdir) / "config.json"
            config_file.write_text("invalid json {{{")

            with patch("lotus123.ui.config.CONFIG_FILE", config_file):
                loaded = AppConfig.load()
                # Should return default config
                assert loaded.theme == "LOTUS"

    def test_save_creates_directory(self):
        """Test save creates config directory if needed."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_dir = Path(tmpdir) / "new_dir"
            config_file = config_dir / "config.json"

            with patch("lotus123.ui.config.CONFIG_DIR", config_dir):
                with patch("lotus123.ui.config.CONFIG_FILE", config_file):
                    config = AppConfig()
                    config.save()
                    assert config_dir.exists()
                    assert config_file.exists()

    def test_save_handles_errors(self):
        """Test save handles write errors gracefully."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create a file instead of directory to cause error
            config_dir = Path(tmpdir) / "config"
            config_dir.write_text("not a directory")

            with patch("lotus123.ui.config.CONFIG_DIR", config_dir):
                config = AppConfig()
                # Should not raise, just silently fail
                config.save()
