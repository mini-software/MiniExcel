# MiniExcel CLI

Query and convert `.xlsx`, `.xlsm`, and `.csv` files from the command line.

## Install

Install the framework-dependent .NET tool:

```shell
dotnet tool install --global MiniExcel.Cli
```

Alternatively, download a NativeAOT archive for Windows, Linux, or macOS from the GitHub release.

## Usage

Convert CSV to Excel or Excel to CSV:

```shell
miniexcel convert --input input.csv --output output.xlsx
miniexcel convert --input input.xlsx --output output.csv --sheet Sheet1 --overwrite
```

Write rows as JSON:

```shell
miniexcel query --input input.xlsx --sheet Sheet1
miniexcel query --input input.csv --start-cell A2 --no-header
```

Run `miniexcel <command> --help` for all options.