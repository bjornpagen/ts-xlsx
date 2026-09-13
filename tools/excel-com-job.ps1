# Excel COM work under a wall-clock watchdog that can only ever kill the Excel it started.
#
# Dot-sourced by the headless COM drivers (tools/excel-oracle/observe.ps1, read-geometry.ps1,
# tools/vba-compiler/compile.ps1). Each once snapshotted the EXCEL.EXE PIDs at start and, on timeout and
# again on the way out, killed every Excel that had appeared since. That is not "the Excel this run
# started": PowerPoint starts Excel instances for its charts at any moment, and two drivers run in
# parallel start each other's, so the sweep could kill a chart someone was editing or a sibling probe.
#
# `New-Object -ComObject Excel.Application` has DCOM start EXCEL.EXE, so the process is a child of
# nothing this run owns. Only its window names it: the work block calls `Register-ExcelProcess` on the
# Application it created, which resolves `Hwnd` to a PID and emits it on the job's output ahead of the
# result, where the watchdog can read it even from a job it had to stop.

function Invoke-ExcelComJob {
  param(
    [Parameter(Mandatory = $true)] [scriptblock] $Work,
    [object[]] $ArgumentList = @(),
    [Parameter(Mandatory = $true)] [int] $TimeoutSec
  )

  $caller = Split-Path -Leaf $MyInvocation.PSCommandPath
  $registration = {
    function Register-ExcelProcess {
      param([Parameter(Mandatory = $true)] $Excel)
      if (-not ('ExcelComJob.User32' -as [type])) {
        Add-Type -Namespace ExcelComJob -Name User32 -MemberDefinition `
          '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);'
      }
      $processId = [uint32]0
      [void][ExcelComJob.User32]::GetWindowThreadProcessId([IntPtr][int64]$Excel.Hwnd, [ref]$processId)
      [pscustomobject]@{ ExcelProcessId = [int]$processId }
    }
  }

  $job = Start-Job -InitializationScript $registration -ScriptBlock $Work -ArgumentList $ArgumentList
  $excelPid = 0
  try {
    $done = Wait-Job -Job $job -Timeout $TimeoutSec
    if ($null -eq $done) { Stop-Job -Job $job -ErrorAction SilentlyContinue }

    # Errors are collected rather than thrown here, so the PID is read before anything unwinds.
    $output = @(Receive-Job -Job $job -ErrorAction SilentlyContinue -ErrorVariable jobErrors)
    $registered = $output | Where-Object { $null -ne $_.PSObject.Properties['ExcelProcessId'] } | Select-Object -First 1
    if ($null -ne $registered) { $excelPid = [int]$registered.ExcelProcessId }

    if ($null -eq $done) {
      # Hung inside COM (a modal that slipped past the guards). Force-kill and fail loudly, never silently.
      $outcome = if (Stop-ExcelProcess $excelPid) {
        "killed the EXCEL.EXE it started (PID $excelPid)"
      } elseif ($excelPid -eq 0) {
        'Excel never reported its process, so nothing was killed and an EXCEL.EXE may be left running'
      } else {
        "the EXCEL.EXE it started (PID $excelPid) had already exited"
      }
      # -ErrorAction Continue: under the callers' 'Stop' preference, Write-Error would throw and exit 1.
      Write-Error "${caller}: Excel COM timed out after ${TimeoutSec}s; $outcome" -ErrorAction Continue
      exit 2
    }

    if ($jobErrors.Count -gt 0) { throw $jobErrors[0] }
    $output | Where-Object { $_ -is [string] }
  } finally {
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    # An Excel that outlived the job's own Quit(): the same process, and still no other.
    [void](Stop-ExcelProcess $excelPid)
  }
}

# Kills the process only while it is still an EXCEL.EXE, so a PID reused since is left alone.
function Stop-ExcelProcess([int] $ProcessId) {
  if ($ProcessId -eq 0) { return $false }
  $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq 'EXCEL' }
  if ($null -eq $process) { return $false }
  $process | Stop-Process -Force -ErrorAction SilentlyContinue
  return $true
}
