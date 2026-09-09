"""Where a frozen SpecCleanse looks for patterns.yaml.

These rules only ever run inside a PyInstaller bundle, which the test suite has
no way to produce, so the bundle is simulated: ``sys.frozen`` and
``sys._MEIPASS`` are what `apppaths` reads to tell the two situations apart.
"""

import os
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import mock

import apppaths
import verify
from detection import DetectionEngine
from docx_xml import load_config
from tests.support import CONFIG_PATH


class FakeBundle:
    """Stand in for a PyInstaller bundle laid out under ``root``.

    ``root/bundle`` is the extraction directory the code and the default
    patterns.yaml are unpacked into; ``root/dist`` holds the executable.
    """

    def __init__(self, root: Path, seed_default: bool = True):
        self.meipass = root / "bundle"
        self.exe_dir = root / "dist"
        self.appdata = root / "appdata"
        self.meipass.mkdir(parents=True, exist_ok=True)
        self.exe_dir.mkdir(parents=True, exist_ok=True)
        if seed_default:
            (self.meipass / "patterns.yaml").write_text(
                "specifier_notes:\n  text_patterns: []\n", encoding="utf-8"
            )

    def patches(self):
        return [
            mock.patch.object(sys, "frozen", True, create=True),
            mock.patch.object(sys, "_MEIPASS", str(self.meipass), create=True),
            mock.patch.object(sys, "executable", str(self.exe_dir / "SpecCleanse.exe")),
            mock.patch.dict(
                os.environ,
                {"APPDATA": str(self.appdata)},
                clear=False,
            ),
        ]


class FrozenLayoutTests(unittest.TestCase):
    def setUp(self):
        self._tmp = TemporaryDirectory()
        # Resolved because Windows hands back an 8.3 short path here
        # (C:/Users/RUNNER~1/...) while apppaths resolves what it returns
        # (C:/Users/runneradmin/...). Same directory, different spelling.
        self.root = Path(self._tmp.name).resolve()
        self.addCleanup(self._tmp.cleanup)

    def _bundle(self, seed_default: bool = True) -> FakeBundle:
        bundle = FakeBundle(self.root, seed_default=seed_default)
        for patch in bundle.patches():
            patch.start()
            self.addCleanup(patch.stop)
        return bundle

    def test_source_checkout_uses_the_file_beside_the_modules(self):
        # Not frozen: the developer layout must be untouched.
        self.assertFalse(apppaths.is_frozen())
        expected = Path(apppaths.__file__).resolve().parent / "patterns.yaml"
        self.assertEqual(apppaths.resolve_config_path(), expected)
        self.assertTrue(expected.is_file(), "the repo's own patterns.yaml")

    def test_a_copy_beside_the_executable_wins(self):
        bundle = self._bundle()
        portable = bundle.exe_dir / "patterns.yaml"
        portable.write_text("specifier_notes:\n", encoding="utf-8")

        self.assertEqual(apppaths.resolve_config_path(), portable)

    def test_first_run_seeds_a_user_copy_from_the_bundled_default(self):
        bundle = self._bundle()
        resolved = apppaths.resolve_config_path()

        self.assertEqual(resolved, bundle.appdata / "SpecCleanse" / "patterns.yaml")
        self.assertTrue(resolved.is_file())
        self.assertEqual(
            resolved.read_text(encoding="utf-8"),
            (bundle.meipass / "patterns.yaml").read_text(encoding="utf-8"),
        )

    def test_an_existing_user_copy_is_never_overwritten(self):
        bundle = self._bundle()
        user_copy = bundle.appdata / "SpecCleanse" / "patterns.yaml"
        user_copy.parent.mkdir(parents=True)
        user_copy.write_text("# edited by the user\n", encoding="utf-8")

        resolved = apppaths.resolve_config_path()

        self.assertEqual(resolved, user_copy)
        self.assertEqual(resolved.read_text(encoding="utf-8"), "# edited by the user\n")

    def test_the_extraction_directory_is_never_the_answer_when_seeding_works(self):
        # PyInstaller deletes _MEIPASS on exit, so edits there vanish silently.
        bundle = self._bundle()
        self.assertNotEqual(
            apppaths.resolve_config_path(), bundle.meipass / "patterns.yaml"
        )

    def test_unwritable_profile_falls_back_to_the_bundled_copy(self):
        bundle = self._bundle()
        with mock.patch.object(
            apppaths.shutil, "copyfile", side_effect=OSError("read-only")
        ):
            resolved = apppaths.resolve_config_path()

        # Degraded but working: shipped patterns, no customisation.
        self.assertEqual(resolved, bundle.meipass / "patterns.yaml")

    def test_resolution_never_raises_when_the_bundle_has_no_default(self):
        # load_config reports a missing file readably; resolution must not blow
        # up first with a traceback no one can see under a windowed build.
        self._bundle(seed_default=False)
        self.assertIsInstance(apppaths.resolve_config_path(), Path)


class VerifierConfigResolutionTests(unittest.TestCase):
    """Which patterns.yaml ``verify_clean`` loads, under each call mode.

    Verification and the cleaner disagreeing about the rules is not a small
    problem: the verifier would judge the output against patterns the cleaner
    never applied and report the difference as damage.  The GUI always supplies
    an engine, so none of this governs a production run today — it is pinned
    because the next engine-less caller will not know that.
    """

    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.root = Path(self._tmp.name).resolve()
        self.addCleanup(self._tmp.cleanup)
        self.loaded: list[Path] = []

        real_load = verify.load_config

        def spy(path):
            self.loaded.append(Path(path))
            return real_load(path)

        patcher = mock.patch.object(verify, "load_config", spy)
        patcher.start()
        self.addCleanup(patcher.stop)

    def _call(self, **kwargs):
        """Run verify_clean far enough to see which config it chose.

        The fixtures do not exist, so it fails once it reaches them; the
        question here is only which file it read on the way.
        """
        with self.assertRaises(Exception):
            verify.verify_clean(self.root / "in.docx", self.root / "out.docx", **kwargs)

    def _bundle(self, seed_default: bool = True) -> FakeBundle:
        bundle = FakeBundle(self.root, seed_default=seed_default)
        for patch in bundle.patches():
            patch.start()
            self.addCleanup(patch.stop)
        return bundle

    def test_an_explicit_path_is_used_as_given(self):
        chosen = self.root / "mine.yaml"
        chosen.write_text("specifier_notes:\n  text_patterns: []\n", encoding="utf-8")

        self._call(config_path=chosen)

        self.assertEqual(self.loaded, [chosen])

    def test_an_engine_means_no_config_is_read_at_all(self):
        # Not merely "the right file": supplying an engine must not resolve a
        # location or seed a per-user copy, because the caller asked for
        # neither.
        engine = DetectionEngine(load_config(CONFIG_PATH))
        self.loaded.clear()

        self._call(engine=engine)

        self.assertEqual(self.loaded, [])

    def test_an_engine_wins_over_an_explicit_path(self):
        # Documented precedence, pinned so it cannot change silently.
        engine = DetectionEngine(load_config(CONFIG_PATH))
        self.loaded.clear()

        self._call(engine=engine, config_path=self.root / "ignored.yaml")

        self.assertEqual(self.loaded, [])

    def test_from_source_the_file_beside_the_modules_is_used(self):
        self._call()

        self.assertEqual(self.loaded, [CONFIG_PATH])

    def test_frozen_the_users_file_wins_over_the_bundled_copy(self):
        # The defect this closed: the old fallback named the module directory,
        # which in a bundle is PyInstaller's extraction directory — temporary,
        # and not where the user's edits live.
        bundle = self._bundle()
        beside_exe = bundle.exe_dir / "patterns.yaml"
        beside_exe.write_text("specifier_notes:\n  text_patterns: []\n", encoding="utf-8")

        self._call()

        self.assertEqual(self.loaded, [beside_exe])
        self.assertNotIn(bundle.meipass / "patterns.yaml", self.loaded)

    def test_frozen_an_explicit_path_still_wins(self):
        self._bundle()
        chosen = self.root / "mine.yaml"
        chosen.write_text("specifier_notes:\n  text_patterns: []\n", encoding="utf-8")

        self._call(config_path=chosen)

        self.assertEqual(self.loaded, [chosen])


if __name__ == "__main__":
    unittest.main()
