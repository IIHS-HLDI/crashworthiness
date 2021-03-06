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
Const DiademExcelWorkbook = "SERXXXXX Summary.xlsm"
Const xlMaximized =&HFFFFEFD7
Dim TESTID 
Dim Test_Type
dim objExcel, objWorkbook 

Test_Type = "Rear"

TESTID = Data.Root.Properties("Test_ID").Value


' Define global variables.
Dim Excel           ' Object variable for Excel
Dim ExcelSheet      ' Object variable for Excel worksheet


' Open Excel files  
Set objExcel = CreateObject("Excel.Application") 
Set objWorkbook = objExcel.Workbooks.Open(CurrentScriptPath & DiademExcelWorkbook) 

' Minimize DIAdem window.
Call WndShow("SHELL", "MINIMIZE")

' Maximize Excel window. 
objExcel.WindowState = xlMaximized 
objExcel.Visible = True 
  
      
  ' Open Rating Sheet
    Set ExcelSheet = objExcel.Workbooks(DiademExcelWorkbook).Sheets("Rating")
   
  'Enter values on the Rating tab
  ExcelSheet.Cells(3,3).Value = Data.Root.Properties("Vehicle_Name").Value
  ExcelSheet.Cells(5,3).Value = Data.Root.Properties("Test_ID").Value
  ExcelSheet.Cells(11,8).Value = Data.Root.Properties("Rated_Backset").Value
  ExcelSheet.Cells(11,9).Value = Data.Root.Properties("Rated_Height").Value
  ExcelSheet.Cells(9,12).Value = Data.Root.Properties("Static_Rating").Value
  ExcelSheet.Cells(17,8).Value = Data.Root.Properties("Maximum_Shear_Force_N_").Value
  ExcelSheet.Cells(19,8).Value = Data.Root.Properties("Maximum_Tension_N_").Value
  ExcelSheet.Cells(17,12).Value = Data.Root.Properties("Neck_Forces_Rating").Value
  ExcelSheet.Cells(22,8).Value = Data.Root.Properties("Maximum_T1_Acceleration_g_").Value
  ExcelSheet.Cells(22,12).Value = Data.Root.Properties("Maximum_T1_Criteria").Value
  ExcelSheet.Cells(24,8).Value = Data.Root.Properties("Head_Contact_Time_ms_").Value
  ExcelSheet.Cells(24,12).Value = Data.Root.Properties("HCT_Criteria").Value
  ExcelSheet.Cells(22,12).Value = Data.Root.Properties("Maximum_T1_Criteria").Value
  ExcelSheet.Cells(13,12).Value = Data.Root.Properties("Dynamic_Rating").Value
  ExcelSheet.Cells(26,12).Value = Data.Root.Properties("Overall_Rating").Value
  'Open Pulse Sheet
    Set ExcelSheet = objExcel.Workbooks(DiademExcelWorkbook).Sheets("Pulse")

  'Enter values on pulse tab
  ExcelSheet.Cells(3,3).Value = Data.Root.Properties("name").Value
  ExcelSheet.Cells(10,10).Value = Data.Root.Properties("t0_Acceleration_g_").Value
  ExcelSheet.Cells(10,11).Value = Data.Root.Properties("t0_Acceleration_Criteria").Value
  ExcelSheet.Cells(12,10).Value = Data.Root.Properties("Peak_Acceleration_g_").Value
  ExcelSheet.Cells(12,11).Value = Data.Root.Properties("Peak_Acceleration_Criteria").Value
  ExcelSheet.Cells(14,10).Value = Data.Root.Properties("End_Time_ms_").Value
  ExcelSheet.Cells(14,11).Value = Data.Root.Properties("End_Time_Criteria").Value
  ExcelSheet.Cells(16,10).Value = Data.Root.Properties("DeltaV_kph_").Value
  ExcelSheet.Cells(16,11).Value = Data.Root.Properties("DeltaV_Criteria").Value

 'Open Channel Comments Tab
    Set ExcelSheet = objExcel.Workbooks(DiademExcelWorkbook).Sheets("Channel Comments")

  'Enter values on Channel Comments Tab
  ExcelSheet.Cells(1,1).Value = Data.Root.Properties("Vehicle_Name").Value
  ExcelSheet.Cells(2,1).Value = Data.Root.Properties("Test_ID").Value

 'Open Summary Tab
    Set ExcelSheet = objExcel.Workbooks(DiademExcelWorkbook).Sheets("Summary")
    
  'Enter values on Summary Tab
  ExcelSheet.Cells(3,2).Value = Data.Root.Properties("Test_ID").Value
  ExcelSheet.Cells(4,2).Value = Data.Root.Properties("Test_Date").Value
  ExcelSheet.Cells(5,2).Value = Data.Root.Properties("HPG_Protocol_Number").Value
  ExcelSheet.Cells(6,2).Value = Data.Root.Properties("ATD_ID").Value
  ExcelSheet.Cells(7,2).Value = Data.Root.Properties("BioRID_Test_Number").Value
  ExcelSheet.Cells(9,2).Value = Data.Root.Properties("ATD_Build_Level").Value
  ExcelSheet.Cells(10,2).Value = Data.Root.Properties("Vehicle_Name").Value
  
  ExcelSheet.Cells(3,4).Value = Data.Root.Properties("Peak_Acceleration_g_").Value
  ExcelSheet.Cells(4,4).Value = Data.Root.Properties("t0_Acceleration_g_").Value
  ExcelSheet.Cells(5,4).Value = Data.Root.Properties("End_Time_ms_").Value
  ExcelSheet.Cells(6,4).Value = Data.Root.Properties("DeltaV_kph_").Value
  ExcelSheet.Cells(7,4).Value = Data.Root.ChannelGroups(2).Channels("S0VEHC0300__ACXD").Properties("Average_Acceleration").Value
  
  ExcelSheet.Cells(21,4).Value = Data.Root.ChannelGroups(2).Channels("S1HEAD0000__ACXA").Properties("Max_Head_Xg").Value
  ExcelSheet.Cells(21,5).Value = (Data.Root.ChannelGroups(2).Channels("S1HEAD0000__ACXA").Properties("Max_Head_Xg_t").Value * 1000)
  ExcelSheet.Cells(22,4).Value = Data.Root.ChannelGroups(3).Channels("Head_Rebound_Velocity").Properties("Max_Reb_Vel").Value
  ExcelSheet.Cells(22,5).Value = (Data.Root.ChannelGroups(3).Channels("Head_Rebound_Velocity").Properties("Max_Reb_Vel_t").Value * 1000)
  ExcelSheet.Cells(23,5).Value = (Data.Root.ChannelGroups(3).Channels("HCT").Properties("HCT_begin").Value * 1000)
  ExcelSheet.Cells(24,5).Value = (Data.Root.ChannelGroups(3).Channels("HCT").Properties("HCT_end").Value * 1000)
  ExcelSheet.Cells(25,4).Value = Data.Root.ChannelGroups(3).Channels("AverageT1XG").Properties("Max_T1Xg").Value
  ExcelSheet.Cells(25,5).Value = (Data.Root.ChannelGroups(3).Channels("AverageT1XG").Properties("Max_T1Xg_t").Value * 1000)
  ExcelSheet.Cells(26,4).Value = Data.Root.ChannelGroups(3).Channels("T1XG_HeadXG").Properties("Max_T1XG_HeadXG").Value
  ExcelSheet.Cells(26,5).Value = (Data.Root.ChannelGroups(3).Channels("T1XG_HeadXG").Properties("Max_T1XG_HeadXG_t").Value * 1000)
  ExcelSheet.Cells(27,4).Value = Data.Root.ChannelGroups(3).Channels("NIC").Properties("Max_NIC").Value
  ExcelSheet.Cells(27,5).Value = (Data.Root.ChannelGroups(3).Channels("NIC").Properties("Max_NIC_t").Value * 1000)
  ExcelSheet.Cells(28,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties("Max_Fx").Value
  ExcelSheet.Cells(28,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties("Max_Fx_t").Value * 1000)
  ExcelSheet.Cells(29,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties("Min_Fx").Value
  ExcelSheet.Cells(29,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties("Min_Fx_t").Value * 1000)
  ExcelSheet.Cells(30,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA").Properties("Tension").Value
  ExcelSheet.Cells(30,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA").Properties("Tension_t").Value * 1000)
  ExcelSheet.Cells(31,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA").Properties("Compression").Value
  ExcelSheet.Cells(31,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA").Properties("Compression_t").Value * 1000)
  ExcelSheet.Cells(32,4).Value = Data.Root.ChannelGroups(3).Channels("MOCy").Properties("Max_My").Value
  ExcelSheet.Cells(32,5).Value = (Data.Root.ChannelGroups(3).Channels("MOCy").Properties("Max_My_t").Value * 1000)
  ExcelSheet.Cells(33,4).Value = Data.Root.ChannelGroups(3).Channels("MOCy").Properties("Min_My").Value
  ExcelSheet.Cells(33,5).Value = (Data.Root.ChannelGroups(3).Channels("MOCy").Properties("Min_My_t").Value * 1000)
  ExcelSheet.Cells(34,4).Value = Data.Root.ChannelGroups(3).Channels("NIJ_NCE").Properties("Max_NCE").Value
  ExcelSheet.Cells(34,5).Value = (Data.Root.ChannelGroups(3).Channels("NIJ_NCE").Properties("Max_NCE_t").Value * 1000)
  ExcelSheet.Cells(35,4).Value = Data.Root.ChannelGroups(3).Channels("NIJ_NCF").Properties("Max_NCF").Value
  ExcelSheet.Cells(35,5).Value = (Data.Root.ChannelGroups(3).Channels("NIJ_NCF").Properties("Max_NCF_t").Value * 1000)
  ExcelSheet.Cells(36,4).Value = Data.Root.ChannelGroups(3).Channels("NIJ_NTE").Properties("Max_NTE").Value
  ExcelSheet.Cells(36,5).Value = (Data.Root.ChannelGroups(3).Channels("NIJ_NTE").Properties("Max_NTE_t").Value * 1000)
  ExcelSheet.Cells(37,4).Value = Data.Root.ChannelGroups(3).Channels("NIJ_NTF").Properties("Max_NTF").Value
  ExcelSheet.Cells(37,5).Value = (Data.Root.ChannelGroups(3).Channels("NIJ_NTF").Properties("Max_NTF_t").Value * 1000)
  ExcelSheet.Cells(38,4).Value = Data.Root.ChannelGroups(3).Channels("Nkm_NEA").Properties("Max_NEA").Value
  ExcelSheet.Cells(38,5).Value = (Data.Root.ChannelGroups(3).Channels("Nkm_NEA").Properties("Max_NEA_t").Value * 1000)
  ExcelSheet.Cells(39,4).Value = Data.Root.ChannelGroups(3).Channels("Nkm_NEP").Properties("Max_NEP").Value
  ExcelSheet.Cells(39,5).Value = (Data.Root.ChannelGroups(3).Channels("Nkm_NEP").Properties("Max_NEP_t").Value * 1000)
  ExcelSheet.Cells(40,4).Value = Data.Root.ChannelGroups(3).Channels("Nkm_NFA").Properties("Max_NFA").Value
  ExcelSheet.Cells(40,5).Value = (Data.Root.ChannelGroups(3).Channels("Nkm_NFA").Properties("Max_NFA_t").Value * 1000)
  ExcelSheet.Cells(41,4).Value = Data.Root.ChannelGroups(3).Channels("Nkm_NFP").Properties("Max_NFP").Value
  ExcelSheet.Cells(41,5).Value = (Data.Root.ChannelGroups(3).Channels("Nkm_NFP").Properties("Max_NFP_t").Value * 1000)
  ExcelSheet.Cells(42,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA").Properties("Max_Lower_Fx").Value
  ExcelSheet.Cells(42,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA").Properties("Max_Lower_Fx_t").Value * 1000)
  ExcelSheet.Cells(43,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA").Properties("Min_Lower_Fx").Value
  ExcelSheet.Cells(43,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOXA").Properties("Min_Lower_Fx_t").Value * 1000)
  ExcelSheet.Cells(44,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA").Properties("Tension").Value
  ExcelSheet.Cells(44,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA").Properties("Tension_t").Value * 1000)
  ExcelSheet.Cells(45,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA").Properties("Compression").Value
  ExcelSheet.Cells(45,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKLO00__FOZA").Properties("Compression_t").Value * 1000)
  ExcelSheet.Cells(46,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB").Properties("Max_Lower_My").Value
  ExcelSheet.Cells(46,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB").Properties("Max_Lower_My_t").Value * 1000)
  ExcelSheet.Cells(47,4).Value = Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB").Properties("Min_Lower_My").Value
  ExcelSheet.Cells(47,5).Value = (Data.Root.ChannelGroups(2).Channels("S1NECKLO00__MOYB").Properties("Min_Lower_My_t").Value * 1000)



Call Run_Excel_TDMS_Addin
objExcel.Windows("SERXXXXX Summary.xlsm").Activate

objExcel.Run ("Move_TDMS_Data")

objExcel.Windows("SERXXXXX Summary.xlsm").Activate

objExcel.ActiveWorkbook.SaveAs "N:\VRC\Shared\Crashworthiness Program\Rear Impact Sled Tests\" & TESTID & "\DATA\EXCEL\" & TESTID & " Summary.xlsm"

' Set object references to 'Nothing'.
Set objExcel = Nothing


' Restore DIAdem window.
Call WndShow("VIEW", "NORMAL")


'Run Excel TDMS Add-in
Sub Run_Excel_TDMS_Addin()

Dim TdmAddIn, TdmsFilePath, Template_Workbook, TDMS_Workbook,TDMS_Name


'objExcel = CreateObject("Excel.Application")
Template_Workbook = objExcel.ActiveWorkbook.Name

TdmsFilePath = "N:\VRC\Shared\Crashworthiness Program\Rear Impact Sled Tests\"

TDMS_NAME = Data.Root.Properties("name").Value


On Error Resume Next

objExcel.Visible = True
If Err.number <> 0 Then Call AutoQuit("Could not contact Excel on this computer")
Set TdmAddIn = objExcel.COMAddIns.Item("ExcelTDM.TdmAddin")
If Err.number <> 0 Then Call AutoQuit("Could not access the TDM Excel Addin" & vbCrLf & vbCrLf & "(TDM Excel Addin version 3.1 or later is required)")
objExcel.Visible = True
Call TdmAddIn.Object.ImportFile(TdmsFilePath & TESTID & "\DATA\DIAdem\" & TDMS_NAME & ".TDMS", True)
If Err.number <> 0 Then Call AutoQuit("Could not load the desired data file" & vbCrLf & vbCrLf & """" & TdmsFilePath & """")
Call objExcel.ActiveWindow.Activate
On Error GoTo 0



End Sub

'---------------------------


