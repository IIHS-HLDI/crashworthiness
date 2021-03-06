' ----------------------------------------------------------------------------------------
' Name         : ---
' Copyright(c) : 
' Author       : 
' Comment      : Create Excel Summary Sheet                                                  
' ----------------------------------------------------------------------------------------

Option Explicit


' Initialize script.
' Suppress display of the command that is currently executing.
AutoEcho = "No"

' ----------------------------------------------------------------------------------------
' MAIN SECTION
' ----------------------------------------------------------------------------------------
' Set constants.
Const PulseExcelWorkbook = "Pulse Summary.xlsm"
Const xlMaximized1 =&HFFFFEFD7
dim objExcel, objWorkbook 
Dim count

count = Data.Root.Properties("HPG_Protocol_Number").Value

' Define global variables.
Dim Excel           ' Object variable for Excel
Dim ExcelSheet      ' Object variable for Excel worksheet


' Open Excel files  
Set objExcel = CreateObject("Excel.Application") 
Set objWorkbook = objExcel.Workbooks.Open("N:\VRC\Shared\Crashworthiness Program\Rear Impact Sled Tests\Pulse Summary.xlsm") 

' Minimize DIAdem window.
Call WndShow("SHELL", "MINIMIZE")

' Maximize Excel window. 
objExcel.WindowState = xlMaximized1 
objExcel.Visible = True 
  

  'Open Pulse Sheet
    Set ExcelSheet = objExcel.Workbooks(PulseExcelWorkbook).Sheets("Pulse Summary")

  'Enter values on pulse summary tab
  ExcelSheet.Cells(count + 2,1).Value = Data.Root.Properties("Test_ID").Value
  ExcelSheet.Cells(count + 2,2).Value = Data.Root.Properties("HPG_Protocol_Number").Value
  ExcelSheet.Cells(count + 2,3).Value = Data.Root.Properties("Test_Date").Value
  ExcelSheet.Cells(count + 2,4).Value = Data.Root.Properties("Peak_Acceleration_g_").Value
  ExcelSheet.Cells(count + 2,5).Value = Data.Root.Properties("Peak_Acceleration_Criteria").Value
  ExcelSheet.Cells(count + 2,6).Value = Data.Root.Properties("DeltaV_kph_").Value
  ExcelSheet.Cells(count + 2,7).Value = Data.Root.Properties("DeltaV_Criteria").Value
  ExcelSheet.Cells(count + 2,8).Value = Data.Root.Properties("t0_Acceleration_g_").Value
  ExcelSheet.Cells(count + 2,9).Value = Data.Root.Properties("t0_Acceleration_Criteria").Value
  ExcelSheet.Cells(count + 2,10).Value = Data.Root.Properties("End_Time_ms_").Value
  ExcelSheet.Cells(count + 2,11).Value = Data.Root.Properties("End_Time_Criteria").Value
  


objExcel.ActiveWorkbook.save
objExcel.ActiveWorkbook.close


' Set object references to 'Nothing'.
Set objExcel = Nothing

    
    

' Restore DIAdem window.
Call WndShow("VIEW", "NORMAL")



'---------------------------


