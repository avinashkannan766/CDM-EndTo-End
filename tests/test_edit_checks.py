
import hashlib
import sys
import unittest
from pathlib import Path

# Locate the project and import the existing edit-check engine.
PROJECT_DIR = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_DIR / "scripts"

sys.path.insert(0, str(SCRIPTS_DIR))

from run_edit_checks import run_checks


CLEAN_FILE = PROJECT_DIR / "data" / "clinical_study_clean.xlsx"
TEST_FILE = PROJECT_DIR / "data" / "clinical_study_test.xlsx"

EXPECTED_CHECK_IDS = {
    "EC-003",
    "EC-004",
    "EC-009",
    "EC-011",
    "EC-017",
    "EC-018",
    "EC-022",
    "EC-025",
}


def file_sha256(path):
    """Return a SHA-256 fingerprint for a file."""
    digest = hashlib.sha256()

    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(8192), b""):
            digest.update(chunk)

    return digest.hexdigest()


class TestEditChecks(unittest.TestCase):

    def test_clean_dataset_has_no_findings(self):
        """The clean synthetic dataset should produce zero findings."""
        self.assertTrue(CLEAN_FILE.exists(), f"Missing file: {CLEAN_FILE}")

        findings = run_checks(CLEAN_FILE)

        self.assertEqual(
            len(findings),
            0,
            "The clean dataset unexpectedly produced findings.",
        )

    def test_test_dataset_has_eight_findings(self):
        """The deliberately seeded test dataset should produce eight findings."""
        self.assertTrue(TEST_FILE.exists(), f"Missing file: {TEST_FILE}")

        findings = run_checks(TEST_FILE)

        self.assertEqual(len(findings), 8)

    def test_test_dataset_contains_expected_check_ids(self):
        """Verify the expected seeded defects were detected."""
        findings = run_checks(TEST_FILE)

        actual_check_ids = {
            finding["Check_ID"]
            for finding in findings
        }

        self.assertEqual(actual_check_ids, EXPECTED_CHECK_IDS)

    def test_test_dataset_severity_counts(self):
        """Verify the expected split of six errors and two review flags."""
        findings = run_checks(TEST_FILE)

        error_count = sum(
            finding["Severity"] == "Error"
            for finding in findings
        )

        review_count = sum(
            finding["Severity"] == "Review"
            for finding in findings
        )

        self.assertEqual(error_count, 6)
        self.assertEqual(review_count, 2)

    def test_finding_ids_are_unique_and_sequential(self):
        """Finding IDs should be unique and repeatable."""
        first_run = run_checks(TEST_FILE)
        second_run = run_checks(TEST_FILE)

        first_ids = [
            finding["Finding_ID"]
            for finding in first_run
        ]
        second_ids = [
            finding["Finding_ID"]
            for finding in second_run
        ]

        expected_ids = [
            f"FND-{number:03d}"
            for number in range(1, 9)
        ]

        self.assertEqual(first_ids, expected_ids)
        self.assertEqual(second_ids, expected_ids)

    def test_source_datasets_are_not_modified(self):
        """Running checks must not change either source workbook."""
        paths = [CLEAN_FILE, TEST_FILE]

        before = {
            path: file_sha256(path)
            for path in paths
        }

        run_checks(CLEAN_FILE)
        run_checks(TEST_FILE)

        after = {
            path: file_sha256(path)
            for path in paths
        }

        self.assertEqual(before, after)


if __name__ == "__main__":
    unittest.main(verbosity=2)
