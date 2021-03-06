'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/27/2017 16:16:15
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")

Class ExcelExport
  Private iPath, oExcel, oWorkbook
  
  Public Property Get Workbook
    Set Workbook = oWorkbook
  End Property
  
  Public Default Function Init(path)
    Call LogError("Export Started...")
    iPath = path
    Set oExcel = CreateObject("Excel.Application")
    Set oWorkbook = oExcel.Workbooks.Open(path) 
    oExcel.WindowState = &HFFFFEFD7
    oExcel.Visible = true 
    pause(3)
    Set Init = Me
  End Function
  
  Private Sub Class_Terminate()
    Call LogError("Export Finished...")
    MsgBox("Export finished with " & LastErrorCount & " errors.")
    AskPrint
    Set oExcel = Nothing
    Set oWorkbook = Nothing
  End Sub
  
  Private Sub AskPrint
    If msgbox("Print summary results?", vbYesNo + vbQuestion) = vbYes Then
      oExcel.Run("Print_Summary")
    End If 
  End Sub
  
  Public Sub DeleteSheet(name)
    oExcel.Application.DisplayAlerts = False
    Call oExcel.Sheets(name).Delete
    oExcel.Application.DisplayAlerts = True
  End Sub
  
  Public Sub SaveAs(path, iihsTestId)
    On Error Resume Next
    Call oWorkbook.SaveAs(path & "\" & iihsTestId & " Summary.xlsm")
    If Err.number <> 0 Then LogError("Summary sheet save failed. Message: " & Replace(Err.Description, vbCrLf, ""))
    On Error Goto 0
  End Sub
  
End Class