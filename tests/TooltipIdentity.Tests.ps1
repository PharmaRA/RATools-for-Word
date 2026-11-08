$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $repo 'scripts\RATools.Build.psm1') -Force
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class RepairExternalWindow {
 [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] struct WndClass {
  public uint style; public IntPtr proc; public int classExtra, windowExtra;
  public IntPtr instance, icon, cursor, background, menu;
  [MarshalAs(UnmanagedType.LPWStr)] public string className;
 }
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern ushort RegisterClass(ref WndClass wc);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern bool UnregisterClass(string name,IntPtr instance);
 [DllImport("kernel32.dll", CharSet=CharSet.Unicode)] static extern IntPtr GetModuleHandle(string name);
 [DllImport("kernel32.dll", CharSet=CharSet.Ansi)] static extern IntPtr GetProcAddress(IntPtr module,string name);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr CreateWindowEx(uint ex,string cls,string title,uint style,int x,int y,int w,int h,IntPtr parent,IntPtr menu,IntPtr instance,IntPtr param);
 [DllImport("user32.dll")] public static extern bool DestroyWindow(IntPtr h);
 const string ClassName="F3 Tooltip RAToolsExternal";
 public static IntPtr Create() {
  var wc=new WndClass {proc=GetProcAddress(GetModuleHandle("user32.dll"),"DefWindowProcW"),instance=GetModuleHandle(null),className=ClassName};
  if(RegisterClass(ref wc)==0) throw new InvalidOperationException("Cannot register fixture class");
  return CreateWindowEx(0x08000080,ClassName,"RATools test external",0x90000000,-30000,-30000,120,30,IntPtr.Zero,IntPtr.Zero,wc.instance,IntPtr.Zero);
 }
 public static void Close(IntPtr h) { DestroyWindow(h); UnregisterClass(ClassName,GetModuleHandle(null)); }
}
'@
$external=[RepairExternalWindow]::Create()
if($external -eq [IntPtr]::Zero){throw 'Cannot create external fixture window'}
$word=$null;$doc=$null
try {
    $word=Start-RAToolsWordSession
    $word.AddIns.Unload($false)
    $doc=$word.Documents.Add()
    $component=$doc.VBProject.VBComponents.Import((Join-Path $repo 'modules\Core_Window.bas'))
    $code=$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
    # Only replace form discovery; filtering, desktop enumeration and SetWindowPos remain real.
    $code=[regex]::Replace($code,'(?ms)^Public Function FindUserFormWindow\(.*?^End Function',@'
Public Function FindUserFormWindow(ByVal windowCaption As String) As LongPtr
    FindUserFormWindow = repairForm
End Function
'@)
    $declarations=@'
Private repairForm As LongPtr
Private Declare PtrSafe Function RepairCreateWindow Lib "user32" Alias "CreateWindowExW" (ByVal ex As Long, ByVal cls As LongPtr, ByVal title As LongPtr, ByVal style As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal parent As LongPtr, ByVal menu As LongPtr, ByVal instance As LongPtr, ByVal param As LongPtr) As LongPtr
Private Declare PtrSafe Function RepairDestroyWindow Lib "user32" Alias "DestroyWindow" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function RepairRegisterClass Lib "user32" Alias "RegisterClassW" (ByRef wc As RepairWindowClass) As Integer
Private Declare PtrSafe Function RepairUnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal cls As LongPtr, ByVal instance As LongPtr) As Long
Private Declare PtrSafe Function RepairModule Lib "kernel32" Alias "GetModuleHandleA" (ByVal name As String) As LongPtr
Private Declare PtrSafe Function RepairProc Lib "kernel32" Alias "GetProcAddress" (ByVal instance As LongPtr, ByVal name As String) As LongPtr
Private Declare PtrSafe Function RepairForeground Lib "user32" Alias "GetForegroundWindow" () As LongPtr
Private Type RepairWindowClass
    style As Long
    proc As LongPtr
    classExtra As Long
    windowExtra As Long
    instance As LongPtr
    icon As LongPtr
    cursor As LongPtr
    background As LongPtr
    menu As LongPtr
    className As LongPtr
End Type
'@
    $code=$code.Replace('Option Explicit',"Option Explicit`r`n"+$declarations)
    $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
    $component.CodeModule.AddFromString($code)
    $probe=@'
Public Function TooltipIdentityProbe() As String
    Dim wc As RepairWindowClass, className As String, plainClass As String
    Dim ordinary As LongPtr, tip As LongPtr, foreground As LongPtr
    Dim message As String
    On Error GoTo Failed
    plainClass = "STATIC"
    className = "F3 Tooltip RAToolsIdentityTest"
    wc.proc = RepairProc(RepairModule("user32.dll"), "DefWindowProcW")
    wc.instance = RepairModule(vbNullString)
    wc.className = StrPtr(className)
    If RepairRegisterClass(wc) = 0 Then Err.Raise 5, , "Fixture class registration failed"
    repairForm = RepairCreateWindow(&H8000080, StrPtr(plainClass), 0, &H90000000, -30000, -30000, 300, 150, 0, 0, wc.instance, 0)
    ordinary = RepairCreateWindow(&H8000080, StrPtr(plainClass), 0, &H90000000, -30000, -30000, 120, 30, 0, 0, wc.instance, 0)
    tip = RepairCreateWindow(&H8000080, StrPtr(className), 0, &H90000000, -30000, -30000, 120, 30, 0, 0, wc.instance, 0)
    If repairForm = 0 Or ordinary = 0 Or tip = 0 Then Err.Raise 5, , "Fixture window creation failed"
    If IsTooltipOfForm(ordinary, repairForm) Then Err.Raise 5, , "Ordinary same-process window was classified as tooltip"
    If IsTooltipOfForm(EXTERNAL_HANDLE, repairForm) Then Err.Raise 5, , "External process window was classified as tooltip"
    If Not IsTooltipOfForm(tip, repairForm) Then Err.Raise 5, , "Valid tooltip identity was rejected"
    foreground = RepairForeground()
    If Not EnsureTooltipOnTop("fixture") Then Err.Raise 5, , "Tooltip raising failed"
    If (GetWindowLongCompat(tip, GWL_EXSTYLE) And WS_EX_TOPMOST) = 0 Then Err.Raise 5, , "Valid tooltip was not raised"
    If (GetWindowLongCompat(ordinary, GWL_EXSTYLE) And WS_EX_TOPMOST) <> 0 Then Err.Raise 5, , "Ordinary window was raised"
    If (GetWindowLongCompat(EXTERNAL_HANDLE, GWL_EXSTYLE) And WS_EX_TOPMOST) <> 0 Then Err.Raise 5, , "External window was raised"
    If RepairForeground() <> foreground Then Err.Raise 5, , "Focus changed"
    message = "PASS"
    GoTo Cleanup
Failed:
    message = Err.Description
Cleanup:
    On Error Resume Next
    If tip <> 0 Then RepairDestroyWindow tip
    If ordinary <> 0 Then RepairDestroyWindow ordinary
    If repairForm <> 0 Then RepairDestroyWindow repairForm
    repairForm = 0
    RepairUnregisterClass StrPtr(className), wc.instance
    TooltipIdentityProbe = message
End Function
'@
    $component.CodeModule.AddFromString($probe.Replace('EXTERNAL_HANDLE',$external.ToInt64().ToString()))
    $result=$word.Run('TooltipIdentityProbe')
    if($result -ne 'PASS'){throw "Tooltip identity failed: $result"}
    Write-Host 'PASS actual window class/process filtering, tooltip raising and focus preservation'
} finally {
    if($null -ne $doc){$doc.Close(0)}
    Stop-RAToolsWordSession -Word $word
    [RepairExternalWindow]::Close($external)
}
Write-Host 'PASS TooltipIdentity.Tests'
