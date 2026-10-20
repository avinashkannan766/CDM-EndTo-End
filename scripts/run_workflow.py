from datetime import datetime
from pathlib import Path
import subprocess
import sys


PROJECT_DIR = Path(__file__).resolve().parent.parent
REPORT_DIR = PROJECT_DIR / "reports"

TEST_RESULTS_FILE = REPORT_DIR / "test_results.txt"
RUN_LOG_FILE = REPORT_DIR / "workflow_run_log.txt"


def run_step(name, command, log_lines):
    """Run one workflow step and record its result."""
    print(f"\n{'=' * 65}")
    print(name)
    print("=" * 65)

    result = subprocess.run(
        command,
        cwd=PROJECT_DIR,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    output = result.stdout + result.stderr

    if output:
        print(output, end="" if output.endswith("\n") else "\n")

    status = "PASS" if result.returncode == 0 else "FAIL"

    log_lines.append(f"\nSTEP: {name}")
    log_lines.append(f"STATUS: {status}")
    log_lines.append(f"COMMAND: {' '.join(command)}")
    log_lines.append(output.rstrip() or "(No console output)")

    if result.returncode != 0:
        log_lines.append(f"EXIT CODE: {result.returncode}")
        print(f"\nWorkflow stopped: {name} failed.")
        return False

    return True


def main():
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    log_lines = [
        "CDM END-TO-END WORKFLOW RUN",
        f"Started: {datetime.now().astimezone().isoformat(timespec='seconds')}",
        "All study records are synthetic.",
        "Existing lifecycle and correction-audit workbooks are preserved.",
    ]

    # Step 1: Run all automated tests.
    test_command = [
        sys.executable,
        "-m",
        "unittest",
        "discover",
        "-s",
        "tests",
        "-p",
        "test_*.py",
        "-v",
    ]

    print("\nRunning automated tests first...")

    test_result = subprocess.run(
        test_command,
        cwd=PROJECT_DIR,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    test_output = test_result.stdout + test_result.stderr

    TEST_RESULTS_FILE.write_text(
        test_output,
        encoding="utf-8",
    )

    print(
        test_output,
        end="" if test_output.endswith("\n") else "\n",
    )

    test_status = (
        "PASS"
        if test_result.returncode == 0
        else "FAIL"
    )

    log_lines.extend([
        "\nSTEP: Automated tests",
        f"STATUS: {test_status}",
        f"COMMAND: {' '.join(test_command)}",
        test_output.rstrip() or "(No test output)",
    ])

    if test_result.returncode != 0:
        log_lines.append(
            "Workflow stopped because automated tests failed."
        )
        log_lines.append(
            f"Finished: "
            f"{datetime.now().astimezone().isoformat(timespec='seconds')}"
        )

        RUN_LOG_FILE.write_text(
            "\n".join(log_lines) + "\n",
            encoding="utf-8",
        )

        print(
            f"\nWorkflow stopped. "
            f"Log saved: {RUN_LOG_FILE}"
        )

        raise SystemExit(1)

    # Step 2 onward: refresh reports in dependency order.
    steps = [
        (
            "Run edit checks",
            [
                sys.executable,
                "scripts/run_edit_checks.py",
            ],
        ),
        (
            "Generate draft queries",
            [
                sys.executable,
                "scripts/generate_queries.py",
            ],
        ),
        (
            "Run reconciliation on the configured input dataset",
            [
                sys.executable,
                "scripts/run_reconciliation.py",
            ],
        ),
        (
            "Build consolidated summary",
            [
                sys.executable,
                "scripts/create_cdm_summary.py",
            ],
        ),
    ]

    for name, command in steps:
        if not run_step(
            name,
            command,
            log_lines,
        ):
            log_lines.append(
                f"Finished: "
                f"{datetime.now().astimezone().isoformat(timespec='seconds')}"
            )

            RUN_LOG_FILE.write_text(
                "\n".join(log_lines) + "\n",
                encoding="utf-8",
            )

            print(
                f"\nWorkflow stopped. "
                f"Log saved: {RUN_LOG_FILE}"
            )

            raise SystemExit(1)

    log_lines.extend([
        "\nOVERALL STATUS: PASS",
        "Note: lifecycle and correction-audit workbooks were not regenerated.",
        f"Finished: "
        f"{datetime.now().astimezone().isoformat(timespec='seconds')}",
    ])

    RUN_LOG_FILE.write_text(
        "\n".join(log_lines) + "\n",
        encoding="utf-8",
    )

    print("\n" + "=" * 65)
    print("WORKFLOW COMPLETED SUCCESSFULLY")
    print(f"Test results: {TEST_RESULTS_FILE}")
    print(f"Workflow log: {RUN_LOG_FILE}")
    print("=" * 65)


if __name__ == "__main__":
    main()