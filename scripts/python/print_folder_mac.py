import os
import subprocess
import time

FOLDER_PATH = "/Users/haitham/Downloads/MyPrintFiles"  # Update to your actual folder

def print_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()

    if ext in [".pdf", ".txt"]:
        print(f"Printing {filepath} via lpr...")
        subprocess.run(["lpr", filepath])

    elif ext in [".xls", ".xlsx"]:
        print(f"Opening Excel file for manual print: {filepath}")
        subprocess.run(["open", "-a", "Microsoft Excel", filepath])
        time.sleep(5)

    else:
        print(f"Skipping unsupported file: {filepath}")


def main():
    for filename in os.listdir(FOLDER_PATH):
        filepath = os.path.join(FOLDER_PATH, filename)
        if os.path.isfile(filepath):
            print_file(filepath)


if __name__ == "__main__":
    main()