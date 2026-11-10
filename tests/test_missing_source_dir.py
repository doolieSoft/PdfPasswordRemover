import os
import sys
sys.path.insert(0, os.path.abspath('.'))

from unittest.mock import patch

import pytest

from PdfPasswordRemover import main


def test_missing_source_dir_fails_cleanly(tmp_path, capsys):
    missing = tmp_path / 'does-not-exist'
    dest = tmp_path / 'out'
    with patch.object(sys, 'argv', ['PdfPasswordRemover', str(missing), str(dest)]):
        with pytest.raises(SystemExit) as exc_info:
            main()
    assert exc_info.value.code == 1
    out = capsys.readouterr().out
    assert 'does not exist' in out
    assert 'Traceback' not in out
