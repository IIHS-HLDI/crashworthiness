'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 09/26/2013 14:57:12
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script. Test comment
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")

Sub OpenTLFExcel (TestID, FrontorSide)
  Dim objExcel, objWorkbook, CrashFolder, fso
  Const xlMaximized = &HFFFFEFD7
  Const TLFWorkbook = "OPENTLF.xlsm"

  Set fso = CreateObject("Scripting.FileSystemObject")

  Call CopyTLF(TestID, FrontorSide)
  
  Dim tlf : tlf = GetTestDirectory(TestID) & "\DATA\DIAdem\" & TestID & ".TLF"
  
  If Not fso.FileExists(tlf) Then
    Exit Sub
  End If
  
  Set objExcel = CreateObject("Excel.Application")
  Set objWorkbook = objExcel.Workbooks.Open(CrashScriptPath & "Misc\" & TLFWorkbook)

  objExcel.WindowState = xlMaximized
  objExcel.Visible = False

  objExcel.Run ("TLFImport(""" & tlf & """)")
  objExcel.Run ("ProcessTLF")

  If FrontorSide = "FRONT" Then
    CrashFolder = "Frontal Crash Tests"
  Else
    CrashFolder = "Side-Impact Crash Tests"
  End If
  
  objExcel.Workbooks(TLFWorkbook).SaveAs "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\" & CrashFolder & "\" & TestID & "\DATA\DIAdem\" & TestID & " TLF.xlsm"
  objExcel.Workbooks(TestID & " TLF.xlsm").Close
  
  set objExcel = Nothing
  set objWorkbook = Nothing
  set fso = Nothing
  
End Sub

Sub CopyTLF (TestID, TestType)
  Dim fso, CrashFolder
  Set fso = CreateObject("Scripting.FileSystemObject")

  If TestType = "FRONT" Then
    CrashFolder = "Frontal Crash Tests"
  Else
    CrashFolder = "Side-Impact Crash Tests"
  End If

  If fso.FolderExists("\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\" & CrashFolder & "\" & TestID & "\DATA\DAS") = True and fso.FileExists("\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\" & CrashFolder & "\" & TestID & "\DATA\DAS\" & TestID & ".TLF") = True Then
    If fso.FolderExists("\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\" & CrashFolder & "\" & TestID) = True Then
      Call fso.CopyFile("\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\" & CrashFolder & "\" & TestID & "\DATA\DAS\" & TestID & ".TLF", "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\" & CrashFolder & "\" & TestID & "\DATA\DIAdem\", True)    
    End If
  End If
  
  Set fso = Nothing
End Sub

Function GetTLFText(TestID)
  dim oExcel, oWorkbook, fso, oSheet, MaxRow, ChanSat, PercentThreshold, Channel, Text, TLFSheet, i
  Const xlDown = -4121
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  If fso.FileExists(GetTestDirectory(TestID) & "\DATA\DIAdem\" & TestID & " TLF.xlsm") = False Then
    Exit Function
  End If
  
  Set oExcel = CreateObject("Excel.Application")
  Set oWorkbook = oExcel.Workbooks.Open(GetTestDirectory(TestID) & "\DATA\DIAdem\" & TestID & " TLF.xlsm")
  
  oExcel.Visible = False
  TLFSheet = "MaxRangevsMaxValue"
  Set oSheet = oExcel.Workbooks(TestID & " TLF.xlsm").Sheets(TLFSheet)
  MaxRow = oSheet.Range("K2").End(xlDown).Row
  Text = ""
  For i = 2 to MaxRow
    ChanSat = oSheet.Range("K" & i).Value
    If ChanSat >= 0.8 Then
      PercentThreshold = Round(ChanSat * 100, 0)
      Channel = oSheet.Range("B" & i).Value
      Text = Text & "-The sensor, " & Channel & ", used " & PercentThreshold & "% of its max range." &  VbCrLf
    End If
  Next
  
  If Text = "" Then
    Text = "No sensors exceeded their max range."
  End If

  oExcel.Workbooks(TestID & " TLF.xlsm").Close
  
  set oExcel = Nothing
  set oWorkbook = Nothing
  set oSheet = Nothing
  
  If fso.FileExists(GetTestDirectory(TestID) & "\DATA\DIAdem\" & TestID & " TLF.xlsm") = True Then
    fso.DeleteFile(GetTestDirectory(TestID) & "\DATA\DIAdem\" & TestID & " TLF.xlsm")
  End If
  
  If fso.FileExists(GetTestDirectory(TestID) & "\DATA\DIAdem\" & TestID & ".TLF") = True Then
    fso.DeleteFile(GetTestDirectory(TestID) & "\DATA\DIAdem\" & TestID & ".TLF")
  End If
  
  set fso = Nothing
  
  GetTLFText = Text
End Function
