# MiniExcel CLI

Query and convert `.xlsx`, `.xlsm`, and `.csv` files from the command line.

## Install

### .NET tool

Install the framework-dependent tool from NuGet:

```shell
dotnet tool install --global MiniExcel.Cli
```

Update or remove it with:

```shell
dotnet tool update --global MiniExcel.Cli
dotnet tool uninstall --global MiniExcel.Cli
```

To test a locally built package:

```shell
dotnet pack src/MiniExcel.Cli/MiniExcel.Cli.csproj --configuration Release --output artifacts/nuget
dotnet tool install MiniExcel.Cli --global --add-source artifacts/nuget
```

### NativeAOT executable

GitHub releases include self-contained archives that do not require a .NET runtime:

- `miniexcel-<version>-win-x64.zip`
- `miniexcel-<version>-linux-x64.tar.gz`
- `miniexcel-<version>-osx-x64.tar.gz`

Extract the archive and place `miniexcel` (or `miniexcel.exe` on Windows) on your `PATH`.

### Use from other languages

The NativeAOT executable can be invoked from any language that can start a process, including Python, Node.js, Rust, Go, Java, and PHP. Use files for input and consume the JSON written by `query` on standard output. No .NET runtime or .NET language binding is required.

```python
import json
import subprocess

result = subprocess.run(
    ["miniexcel", "query", "--input", "input.xlsx"],
    check=True,
    capture_output=True,
    text=True,
)
rows = json.loads(result.stdout)
```

This is command-line process integration, not a native library ABI or direct FFI API.

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

## Build and publish

Build and test locally:

```shell
dotnet test tests/MiniExcel.Cli.Tests/MiniExcel.Cli.Tests.csproj --configuration Release
dotnet publish src/MiniExcel.Cli/MiniExcel.Cli.csproj --configuration Release --runtime win-x64
```

Pushing a semantic version tag such as `1.46.1` triggers the CLI release workflow. It builds and smoke-tests NativeAOT executables for Windows, Linux, and macOS, creates the .NET tool package, and attaches all artifacts to the GitHub release.