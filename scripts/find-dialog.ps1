# Prints 1 if process <pid> owns a visible standard Windows dialog (#32770), else 0.
param([int]$ProcessId)
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public class Win32Probe {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
}
'@
$script:found = 0
[Win32Probe]::EnumWindows({
    param($h, $l)
    $procId = 0
    [Win32Probe]::GetWindowThreadProcessId($h, [ref]$procId) | Out-Null
    if ($procId -eq $ProcessId -and [Win32Probe]::IsWindowVisible($h)) {
      $sb = New-Object System.Text.StringBuilder 64
      [Win32Probe]::GetClassName($h, $sb, 64) | Out-Null
      if ($sb.ToString() -eq '#32770') { $script:found = 1 }
    }
    return $true
  }, [IntPtr]::Zero) | Out-Null
Write-Output $script:found
