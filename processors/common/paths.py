from pathlib import Path


DATA_SUBDIRECTORIES = (
    "reference_number_supplement",
    "receipt_statistics",
    "subsidy_coupon_statistics",
    "submitted",
)


def choose_existing_data_dir() -> Path | None:
    print(r"Enter the data directory path (for example, C:\Users\username\data):")
    while True:
        try:
            value = input("Data directory: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nProcessing cancelled")
            return None

        if not value:
            print("The path cannot be empty. Try again.")
            continue

        data_dir = Path(value.strip('"')).expanduser().resolve()
        if data_dir.is_dir():
            return data_dir
        print(f"Directory does not exist: {data_dir}")


def resolve_data_dir() -> Path | None:
    current_data_dir = (Path.cwd() / "data").resolve()
    if current_data_dir.is_dir():
        print(f"Found data directory: {current_data_dir}")
        return current_data_dir

    print(f"No data directory found in the current directory: {current_data_dir}")
    while True:
        print("Choose how to configure the data directory:")
        print("  1. Use an existing data directory")
        print("  2. Create the standard data directory structure here")
        print("  0. Exit")
        try:
            choice = input("Enter a number: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nProcessing cancelled")
            return None

        if choice == "0":
            print("Exited")
            return None
        if choice == "1":
            return choose_existing_data_dir()
        if choice == "2":
            if current_data_dir.exists():
                print(f"Cannot create directory; the path is a file: {current_data_dir}")
                continue
            current_data_dir.mkdir()
            for directory_name in DATA_SUBDIRECTORIES:
                (current_data_dir / directory_name).mkdir()
            print(f"Created data directory structure: {current_data_dir}")
            print("Add the required files to each subdirectory, then run the program again.")
            print("This run has ended.")
            return None

        print("Invalid input. Enter a menu number.")


def resolve_existing_data_file(data_dir: Path, candidates: tuple[Path, ...]) -> Path:
    for candidate in candidates:
        path = data_dir / candidate
        if path.exists():
            return path
    return data_dir / candidates[0]
