from pathlib import Path


def main() -> None:
    files: list = [Path]
    text_file: Path

    curr_path = Path.cwd()
    if curr_path.name.find('utils'):
        files = sorted
    files = sorted([f.name for f in Path.cwd().parent.glob("outputs/*.txt")])

    for text_file in files:
        try:
            with open(f"../outputs/{text_file}", 'r') as f:
                text = f.read()
                print(f"slide:{text_file}")
                print("-----")
                print(f"{text}")
                print("-----")

                input('any key for next slide')
        except FileNotFoundError:
            print("FileNotFoundError")


if __name__ == "__main__":
    main()
