param(
    [string]$Workbook,
    [string]$RustRepository,
    [ValidateRange(1, 100)]
    [int]$Passes = 3,
    [ValidateRange(1, 100)]
    [int]$Iterations = 3,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$repositoryRoot = Split-Path $PSScriptRoot -Parent

if ([string]::IsNullOrWhiteSpace($Workbook)) {
    $benchmarkData = Join-Path $PSScriptRoot "MiniExcel.Benchmarks" "data"
    $Workbook = Join-Path $benchmarkData "Test100,000x10.xlsx"
}
if ([string]::IsNullOrWhiteSpace($RustRepository)) {
    $RustRepository = Join-Path (Split-Path $repositoryRoot -Parent) "MiniExcel-Rust"
}

$Workbook = [IO.Path]::GetFullPath($Workbook)
$RustRepository = [IO.Path]::GetFullPath($RustRepository)
$dotnetStressDirectory = Join-Path $PSScriptRoot "MiniExcel.StressTests"
$dotnetProject = Join-Path $dotnetStressDirectory "MiniExcel.StressTests.csproj"
$dotnetOutputDirectory = Join-Path $dotnetStressDirectory "bin" "Release" "net10.0"
$dotnetRunner = Join-Path $dotnetOutputDirectory "MiniExcel.StressTests.dll"
$rustManifest = Join-Path $RustRepository "Cargo.toml"
$executableSuffix = if ($env:OS -eq "Windows_NT") { ".exe" } else { "" }
$rustOutputDirectory = Join-Path $RustRepository "target" "release" "examples"
$rustRunner = Join-Path $rustOutputDirectory "stress_query$executableSuffix"

foreach ($requiredPath in @($Workbook, $dotnetProject, $rustManifest)) {
    if (-not (Test-Path $requiredPath)) {
        throw "Required path not found: $requiredPath"
    }
}

if (-not $SkipBuild) {
    & dotnet build $dotnetProject -c Release --nologo --verbosity:quiet
    if ($LASTEXITCODE -ne 0) { throw ".NET stress runner build failed." }

    & cargo +1.85.0 build --manifest-path $rustManifest --release -p miniexcel --example stress_query --locked
    if ($LASTEXITCODE -ne 0) { throw "Rust stress runner build failed." }
}

function Invoke-MeasuredProcess {
    param(
        [string]$Runtime,
        [string]$Executable,
        [string[]]$Arguments,
        [int]$Iteration
    )

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $Executable
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in $Arguments) {
        $startInfo.ArgumentList.Add($argument)
    }

    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $process = [Diagnostics.Process]::Start($startInfo)
    $peakWorkingSet = 0L
    while (-not $process.WaitForExit(10)) {
        $process.Refresh()
        $peakWorkingSet = [Math]::Max($peakWorkingSet, $process.WorkingSet64)
    }
    $stopwatch.Stop()

    $standardOutput = $process.StandardOutput.ReadToEnd().Trim()
    $standardError = $process.StandardError.ReadToEnd().Trim()
    if ($process.ExitCode -ne 0) {
        throw "$Runtime runner failed with exit code $($process.ExitCode): $standardError"
    }

    $rowCount = 0L
    if (-not [long]::TryParse($standardOutput, [ref]$rowCount)) {
        throw "$Runtime runner returned an invalid row count: $standardOutput"
    }

    [pscustomobject]@{
        Runtime = $Runtime
        Iteration = $Iteration
        Passes = $Passes
        Rows = $rowCount
        ElapsedMs = [Math]::Round($stopwatch.Elapsed.TotalMilliseconds, 2)
        PeakWorkingSetMB = [Math]::Round($peakWorkingSet / 1MB, 2)
    }
}

$runners = @(
    @{
        Runtime = ".NET"
        Executable = "dotnet"
        Arguments = @($dotnetRunner, $Workbook, $Passes.ToString())
    },
    @{
        Runtime = "Rust"
        Executable = $rustRunner
        Arguments = @($Workbook, $Passes.ToString())
    }
)

foreach ($runner in $runners) {
    $null = Invoke-MeasuredProcess @runner -Iteration 0
}

$results = foreach ($iteration in 1..$Iterations) {
    foreach ($runner in $runners) {
        Invoke-MeasuredProcess @runner -Iteration $iteration
    }
}

$expectedRows = $results[0].Rows
if ($results.Where({ $_.Rows -ne $expectedRows }).Count -ne 0) {
    throw "The runners returned different row counts."
}

$results | Format-Table -AutoSize

$summary = $results | Group-Object Runtime | ForEach-Object {
    $elapsed = $_.Group.ElapsedMs | Measure-Object -Average
    $memory = $_.Group.PeakWorkingSetMB | Measure-Object -Average -Maximum
    [pscustomobject]@{
        Runtime = $_.Name
        AverageElapsedMs = [Math]::Round($elapsed.Average, 2)
        AveragePeakWorkingSetMB = [Math]::Round($memory.Average, 2)
        MaximumPeakWorkingSetMB = [Math]::Round($memory.Maximum, 2)
        RowsPerIteration = $expectedRows
    }
}

"Summary"
$summary | Format-Table -AutoSize