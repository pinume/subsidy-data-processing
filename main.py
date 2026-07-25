import os
import sys
import traceback
from types import ModuleType

from processors import digital, large_appliances
from processors.common.paths import resolve_data_dir


PROJECTS: tuple[tuple[str, ModuleType], ...] = (
    ("large appliances", large_appliances),
    ("digital", digital),
)


def choose_project() -> ModuleType | None:
    print("Select a data type:")
    for index, (project_name, _) in enumerate(PROJECTS, start=1):
        print(f"  {index}. {project_name}")
    print("  0. Exit")

    while True:
        try:
            choice = input("Enter a number: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nProcessing cancelled")
            return None

        if choice == "0":
            print("Exited")
            return None

        for index, (project_name, project) in enumerate(PROJECTS, start=1):
            if choice == str(index) or choice == project_name:
                print(f"Selected: {project_name}")
                return project

        print("Invalid input. Enter a menu number or data type name.")


def report_failure(error: BaseException) -> None:
    """Show operators the cause, not a Python stack trace."""
    print(f"\nProcessing failed: {error}", file=sys.stderr)
    print(
        "Existing output files were left unchanged. "
        "Check the source files named above, then run the program again.",
        file=sys.stderr,
    )
    if os.environ.get("UPLOAD_DATA_DEBUG"):
        traceback.print_exc()
    else:
        print(
            "Set UPLOAD_DATA_DEBUG=1 for the full traceback.",
            file=sys.stderr,
        )


def main() -> int:
    data_dir = resolve_data_dir()
    if data_dir is None:
        return 0

    project = choose_project()
    if project is None:
        return 0

    project.configure_data_dir(data_dir)
    processor = project.choose_data_processor()
    if processor is None:
        return 0

    try:
        processor()
    except KeyboardInterrupt:
        print("\nProcessing cancelled", file=sys.stderr)
        return 130
    except Exception as error:
        report_failure(error)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
