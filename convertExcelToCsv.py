import pandas as pd
from rich.markdown import Markdown
from rich.console import Console

import contextlib

from rich.errors import NotRenderableError

# from rich import print
from rich.table import Table
from rich.text import Text

import argparse
import os


def rich_display_dataframe(df, title="Dataframe") -> None:
    # ensure dataframe contains only string values
    df = df.astype(str)

    table = Table(title=title)
    for col in df.columns:
        table.add_column(col)
    for row in df.values:
        with contextlib.suppress(NotRenderableError):
            table.add_row(*row)
    console.print(table)


def get_file_name_without_extension(file_path):
    # Split the file path into a pair (root, ext)
    file_name_with_extension = os.path.basename(file_path)
    file_name, _ = os.path.splitext(file_name_with_extension)
    return file_name


def read_file(file_path):
    try:
        # Try to read the file as an Excel file
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Reading as Excel failed with error: {e}. Trying as TIS-620")
        df = pd.read_excel(file_path, encoding="TIS-620")
    return df


def convertExcelToCsv(file):
    df.to_csv(
        os.path.join(
            os.path.dirname(file), get_file_name_without_extension(file) + ".csv"
        ),
        index=False,
    )


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Excel to CSV converter")
    parser.add_argument("file", type=str, help="Excel file path")
    # parser.add_argument("--csv", action="store_true", help="Convert excel to csv")
    parser.add_argument("--show", action="store_true", help="Print the dataframe")
    args = parser.parse_args()

    console = Console()
    df = read_file(args.file)

    if args.show:
        rich_display_dataframe(df)
    else:
        convertExcelToCsv(args.file)
        console.print(
            Text(
                f"=== [Completed] === Saved as {os.path.dirname(args.file) + get_file_name_without_extension(args.file) + '.csv'}",
                style="bold blue",
            )
        )
