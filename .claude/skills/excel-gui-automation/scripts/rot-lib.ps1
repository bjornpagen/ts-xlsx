<#
.SYNOPSIS
  Dot-source library: reach a workbook some running Excel holds open through the
  Running Object Table, under pwsh 7 as under Windows PowerShell.

.DESCRIPTION
  [Runtime.InteropServices.Marshal]::GetActiveObject does not exist on .NET Core,
  so it throws under pwsh 7. It also named the wrong instance whenever more than
  one EXCEL.EXE ran: it returns the Application registered first, not the /x
  instance open-verdict.ps1 spawned.

  Excel registers every workbook it holds open in the ROT under a moniker whose
  display name is the workbook's full path. Enumerating the ROT and binding the
  matching entry reaches that workbook in whichever instance holds it; its
  .Application is that instance. Marshal.BindToMoniker is not used: bound to a
  path no Excel holds open, a file moniker opens the file, in a new hidden Excel.
#>

if (-not ('XlRot' -as [type])) {
  # No -ReferencedAssemblies: passing it replaces the default references rather than adding to them.
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class XlRot {
  [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);
  [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved, out IBindCtx ctx);

  public static string[] DisplayNames() {
    var names = new string[0];
    Walk((name, rot, moniker) => {
      Array.Resize(ref names, names.Length + 1);
      names[names.Length - 1] = name;
      return false;
    });
    return names;
  }

  // The object registered under exactly this display name, or null.
  public static object Get(string displayName) {
    object found = null;
    Walk((name, rot, moniker) => {
      if (!string.Equals(name, displayName, StringComparison.OrdinalIgnoreCase)) return false;
      rot.GetObject(moniker, out found);
      return true;
    });
    return found;
  }

  static void Walk(Func<string, IRunningObjectTable, IMoniker, bool> visit) {
    IRunningObjectTable rot;
    IBindCtx ctx;
    Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out rot));
    Marshal.ThrowExceptionForHR(CreateBindCtx(0, out ctx));
    IEnumMoniker monikers;
    rot.EnumRunning(out monikers);
    var one = new IMoniker[1];
    while (monikers.Next(1, one, IntPtr.Zero) == 0) {
      string name;
      try { one[0].GetDisplayName(ctx, null, out name); } catch (COMException) { continue; }
      if (visit(name, rot, one[0])) return;
    }
  }
}
'@
}

# The open workbook whose ROT display name, its full path, or that path's leaf matches the wildcard,
# from whichever Excel holds it; $null when none does.
function Get-RunningWorkbook {
  param([Parameter(Mandatory)][string]$NameLike)
  foreach ($name in [XlRot]::DisplayNames()) {
    if ($name -like $NameLike -or [IO.Path]::GetFileName($name) -like $NameLike) {
      return [XlRot]::Get($name)
    }
  }
  return $null
}

# Run a COM call against a running Excel, retrying while Excel refuses calls because it is busy: just
# after a repair it answers RPC_E_CALL_REJECTED (0x80010001) or RPC_E_SERVERCALL_RETRYLATER (0x8001010A)
# for a few seconds. Any other failure, or the last busy one, is rethrown.
function Invoke-WhenExcelAccepts {
  param([Parameter(Mandatory)][scriptblock]$Call, [int]$TimeoutSec = 30)
  $deadline = (Get-Date).AddSeconds($TimeoutSec)
  while ($true) {
    try { return & $Call } catch {
      $hresult = $_.Exception.HResult
      if (-not $_.Exception.InnerException) { $inner = $_.Exception } else { $inner = $_.Exception.InnerException }
      $busy = @(0x80010001, 0x8001010A) | Where-Object { $hresult -eq [int]$_ -or $inner.HResult -eq [int]$_ }
      if (-not $busy -or (Get-Date) -ge $deadline) { throw }
      Start-Sleep -Milliseconds 500
    }
  }
}

# Every display name the ROT holds, for diagnosing a workbook Get-RunningWorkbook did not find.
function Get-RunningObjectNames { [XlRot]::DisplayNames() }
