# MiniExcel Query memory validation

This project enumerates the same `1,000,000 x 10` workbook used by the main
benchmark project. It reports cumulative managed allocations separately from
peak process memory:

- `Total managed allocated` is the total allocated over the complete query. It
  is comparable to BenchmarkDotNet's `Allocated` column, but it is not memory
  held at one time.
- `Peak working set` and `Peak private bytes` are sampled from a separate parent
  process while the query worker runs.

Run against the current MiniExcel source:

```powershell
dotnet run --project benchmarks/MiniExcel.MemoryValidation -c Release
```

Run against a released NuGet package, such as the version in the comparison:

```powershell
dotnet run --project benchmarks/MiniExcel.MemoryValidation -c Release -p:MiniExcelVersion=1.45.0
```

Use another workbook with `--file`:

```powershell
dotnet run --project benchmarks/MiniExcel.MemoryValidation -c Release -- --file C:\data\input.xlsx
```

Use a Release build and compare runs on the same machine and runtime. The
default target is .NET 10 so it matches the local comparison environment.