import contextlib
import io
import os
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from pypdf import PdfReader, PdfWriter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import PdfPasswordRemover as ppr  # noqa: E402


def make_pdf(path, owner_password=None, user_password=''):
    """Write a one-page PDF, optionally protected by an owner password."""
    writer = PdfWriter()
    writer.add_blank_page(width=72, height=72)
    if owner_password is not None:
        writer.encrypt(user_password=user_password, owner_password=owner_password, algorithm='AES-256')
    with open(path, 'wb') as f:
        writer.write(f)
    return path


def make_temp_dir(test_case):
    tmp = tempfile.TemporaryDirectory()
    test_case.addCleanup(tmp.cleanup)
    return Path(tmp.name)


def run_main(*argv):
    """Run main() and return (exit_code, stdout, stderr)."""
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = ppr.main([str(a) for a in argv])
    return code, out.getvalue(), err.getvalue()


class IsPdfTest(unittest.TestCase):
    def test_accepted_extensions(self):
        for name in ('a.pdf', 'a.PDF', 'a.pdf.convert', 'a.PDF.Import'):
            with self.subTest(name=name):
                self.assertTrue(ppr.is_pdf(Path(name)))

    def test_rejected_extensions(self):
        for name in ('a.txt', 'a.pdfx', 'a.convert', 'pdf', 'a.pdf.bak'):
            with self.subTest(name=name):
                self.assertFalse(ppr.is_pdf(Path(name)))


class OutputNameTest(unittest.TestCase):
    def test_plain_pdf_is_unchanged(self):
        self.assertEqual(ppr.output_name('Report.PDF'), 'Report.PDF')

    def test_extra_suffix_is_stripped_keeping_case(self):
        self.assertEqual(ppr.output_name('Report.PDF.convert'), 'Report.PDF')
        self.assertEqual(ppr.output_name('Report.pdf.IMPORT'), 'Report.pdf')

    def test_suffix_not_after_pdf_is_kept(self):
        self.assertEqual(ppr.output_name('file.convert'), 'file.convert')


class CheckDirectoriesTest(unittest.TestCase):
    def setUp(self):
        self.tmp = make_temp_dir(self)

    def test_both_exist(self):
        self.assertIsNone(ppr.check_directories(self.tmp, self.tmp))

    def test_missing_source(self):
        error = ppr.check_directories(self.tmp / 'missing', self.tmp)
        self.assertIn('Source directory', error)

    def test_missing_destination(self):
        error = ppr.check_directories(self.tmp, self.tmp / 'missing')
        self.assertIn('Destination directory', error)

    def test_destination_is_a_file(self):
        file_path = self.tmp / 'file'
        file_path.touch()
        error = ppr.check_directories(self.tmp, file_path)
        self.assertIn('Destination directory', error)


class MainTest(unittest.TestCase):
    def setUp(self):
        tmp = make_temp_dir(self)
        self.source = tmp / 'source'
        self.destination = tmp / 'destination'
        self.source.mkdir()
        self.destination.mkdir()

    def created_dirs(self, kind):
        dirs = list(self.destination.glob(f'*_{kind}'))
        self.assertEqual(len(dirs), 1, f"expected one *_{kind} directory, got {dirs}")
        return dirs[0]

    def test_missing_destination_exits_cleanly(self):
        pdf = make_pdf(self.source / 'a.pdf', owner_password='secret')
        missing = self.destination / 'missing'

        code, _, err = run_main(self.source, missing)

        self.assertEqual(code, 1)
        self.assertIn('Destination directory does not exist', err)
        self.assertFalse(missing.exists())
        self.assertTrue(pdf.exists())

    def test_missing_source_exits_cleanly(self):
        code, _, err = run_main(self.source / 'missing', self.destination)

        self.assertEqual(code, 1)
        self.assertIn('Source directory does not exist', err)
        self.assertEqual(list(self.destination.iterdir()), [])

    def test_missing_directory_still_pauses_when_requested(self):
        with mock.patch('builtins.input') as fake_input:
            code, _, _ = run_main(self.source, self.destination / 'missing', '--pause')
        self.assertEqual(code, 1)
        fake_input.assert_called_once()

    def test_protected_pdf_is_unlocked_and_original_backed_up(self):
        make_pdf(self.source / 'a.pdf', owner_password='secret')
        self.assertTrue(PdfReader(self.source / 'a.pdf').is_encrypted)

        code, out, _ = run_main(self.source, self.destination)

        self.assertEqual(code, 0)
        self.assertIn('1 pdf file(s) have been treated', out)
        new_file = self.created_dirs('new') / 'a.pdf'
        self.assertFalse(PdfReader(new_file).is_encrypted)
        self.assertEqual(len(PdfReader(new_file).pages), 1)
        self.assertTrue((self.created_dirs('backup') / 'a.pdf').exists())
        self.assertFalse((self.source / 'a.pdf').exists())

    def test_extra_suffix_is_removed_from_output(self):
        make_pdf(self.source / 'a.pdf.convert', owner_password='secret')

        run_main(self.source, self.destination)

        self.assertTrue((self.created_dirs('new') / 'a.pdf').exists())
        self.assertTrue((self.created_dirs('backup') / 'a.pdf.convert').exists())

    def test_non_pdf_files_are_ignored(self):
        other = self.source / 'notes.txt'
        other.write_text('hello')

        code, out, _ = run_main(self.source, self.destination)

        self.assertEqual(code, 0)
        self.assertIn('0 pdf file(s)', out)
        self.assertTrue(other.exists())
        self.assertEqual(list(self.destination.iterdir()), [])

    def test_invalid_pdf_is_left_in_place(self):
        broken = self.source / 'broken.pdf'
        broken.write_bytes(b'not a pdf')

        code, out, _ = run_main(self.source, self.destination)

        self.assertEqual(code, 0)
        self.assertIn('1 pdf file(s) could not be treated', out)
        self.assertTrue(broken.exists())
        self.assertEqual(list(self.created_dirs('new').iterdir()), [])

    def test_pdf_requiring_user_password_is_left_in_place(self):
        locked = make_pdf(self.source / 'locked.pdf', owner_password='owner', user_password='user')

        _, out, _ = run_main(self.source, self.destination)

        self.assertIn('1 pdf file(s) could not be treated', out)
        self.assertTrue(locked.exists())

    def test_young_files_are_skipped(self):
        old = make_pdf(self.source / 'old.pdf', owner_password='secret')
        young = make_pdf(self.source / 'young.pdf', owner_password='secret')
        an_hour_ago = time.time() - 3600
        os.utime(old, (an_hour_ago, an_hour_ago))

        run_main(self.source, self.destination, '-a', '60')

        self.assertTrue(young.exists())
        self.assertFalse(old.exists())
        self.assertEqual([p.name for p in self.created_dirs('new').iterdir()], ['old.pdf'])


if __name__ == '__main__':
    unittest.main()
