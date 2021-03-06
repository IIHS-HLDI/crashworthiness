' ----------------------------------------------------------------------------------------

' Author       : 
' Comment      : ---                                                  
' ----------------------------------------------------------------------------------------

Option Explicit
  'Pull in crash tools script
  Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")
  
  'Begin Log
  Call LogError("Export Started...")

  'Suppress display of the command that is currently executing.
  AutoEcho = "No"

  'Set constants.
  Const DiademExcelWorkbook = "CESXXXX Summary.xlsm"
  Const xlMaximized =&HFFFFEFD7

  'Initialize variables
  Dim testId : testId = Data.Root.Properties("Test_ID").Value
  Dim driverDataExist : driverDataExist = Data.Root.ChannelGroups.Exists("11_Head_Neck")
  Dim passengerDataExist : passengerDataExist = Data.Root.ChannelGroups.Exists("14_Head_Neck")
  
  'Open Excel files  
  Dim objExcel : Set objExcel = CreateObject("Excel.Application") 
  Dim objWorkbook : Set objWorkbook = objExcel.Workbooks.Open(CurrentScriptPath & DiademExcelWorkbook) 

  'Minimize DIAdem window.
  Call WndShow("SHELL", "MINIMIZE")

  'Maximize Excel window. 
  objExcel.WindowState = xlMaximized 
  objExcel.Visible = True 

  If driverDataExist Then     
    Call Transfer_Peak_Measures(11, objExcel.Workbooks(DiademExcelWorkbook).Sheets("Driver Summary Sheet"))
    'Run macros in the Excel file 
    objExcel.Run ("Highlight") 
    objExcel.Workbooks(DiademExcelWorkbook).Sheets("Driver Summary Sheet").Name = TESTID + " Driver Summary"
  End If

  If passengerDataExist Then    
    Call Transfer_Peak_Measures(14, objExcel.Workbooks(DiademExcelWorkbook).Sheets("Passenger Summary Sheet"))
    'Run macros in the Excel file 
    objExcel.Run ("Highlight") 
    objExcel.Workbooks(DiademExcelWorkbook).Sheets("Passenger Summary Sheet").Name = TESTID + " Passenger Summary"
  End If

  Call FilmAnalysisExport(driverDataExist, passengerDataExist)
  
  objExcel.Windows("CESXXXX Summary.xlsm").Activate
  
  'End Log
  Call LogError("Export Finished...")
  MsgBox("Export finished with " & LastErrorCount & " errors.")
  Call LaunchLog
  
  objExcel.ActiveWorkbook.SaveAs "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\Side-Impact Crash Tests\" & TESTID & "\DATA\EXCEL\" & TESTID & " Summary.xlsm"
  
  If msgbox("Print summary results?", vbYesNo + vbQuestion) = vbYes Then
    objExcel.Run("Print_Summary")
  End If 
  
  'Set object references to 'Nothing'.
  Set objExcel = Nothing
  Set objWorkbook = Nothing

  'Restore DIAdem window.
  Call WndShow("SHELL", "NORMAL")
  
  'This sub transfers header and peak injury measures to the summary sheet in the excel document
  Sub Transfer_Peak_Measures(Occ_Position, ExcelSheet)
    ExcelSheet.Select
    'Transfer header information
    ExcelSheet.Cells(1, 1).Select
    Call TransferPropToCellRef(ExcelSheet, 1, 1, eSearchFile, "", "", "Test_ID", "", "")
    Call TransferPropToCellRef(ExcelSheet, 1, 2, eSearchFile, "", "", "Vehicle_Name", "", "")
    Call TransferPropToCellRef(ExcelSheet, 2, 2, eSearchFile, "", "", "Description", "", "")
    Call TransferPropToCellRef(ExcelSheet, 3, 2, eSearchFile, "", "", "Impact_velocity", 1, 1)
    Call TransferPropToCellRef(ExcelSheet, 1, 6, eSearchFile, "", "", "Test_Date", "", "")
    ExcelSheet.Cells(2,6).Value = date()
  
    Dim i, j, temp_index  'This will be the first row and column to put data into (j should eventually = 4)
    
    i = 4  'row
    j = 4  'column
  
    temp_index = Data.Root.ChannelGroups(Occ_Position & "_Head_Neck").Properties("index").Value
    
    'Resultant head acceleration
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "HEADCG00__ACR", "Maximum", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "HEADCG00__ACR", "Max_time", 0, 1)    
    i = i + 1
    
    'Resultant head acceleration 3 ms clip
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "HEADCG00__ACR", "Clip_3_ms", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "HEADCG00__ACR", "Clip_3_ms_T1", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+2, eSearchChannel, temp_index, Occ_Position & "HEADCG00__ACR", "Clip_3_ms_T2", 0, 1)   
    i = i + 1

    'HIC15
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "HEADCG00__ACR", "HIC15", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "HEADCG00__ACR", "HIC15_T1", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+2, eSearchChannel, temp_index, Occ_Position & "HEADCG00__ACR", "HIC15_T2", 0, 1)
    i = i + 1

    'Neck A-P Shear force
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "NECKUP00__FOXA", "Max_force", 1, 1000)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "NECKUP00__FOXA", "Max_time", 0, 1)    
    i = i + 1
    
    'Neck L-M Shear force  
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "NECKUP00__FOYA", "Max_force", 1, 1000)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "NECKUP00__FOYA", "Max_time", 0, 1)    
    i = i + 1
    
    'Neck tension
    If Data.Root.ChannelGroups(temp_index).Channels.Exists(Occ_Position & "NECKUP00__FOZA") Then 
      If Data.Root.ChannelGroups(temp_index).Channels(Occ_Position & "NECKUP00__FOZA").Properties("Maximum").Value < 0 Then
        ExcelSheet.Cells(i, j).Value = 0
        ExcelSheet.Cells(i, j+1).Value = 0
      Else
        Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "NECKUP00__FOZA", "Maximum", 1, 1000)
        Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "NECKUP00__FOZA", "Max_time", 0, 1)    
      End If
      i = i + 1
      
      'Neck compression
      If Data.Root.ChannelGroups(temp_index).Channels(Occ_Position & "NECKUP00__FOZA").Properties("Minimum").Value > 0 Then
        ExcelSheet.Cells(i, j).Value = 0
        ExcelSheet.Cells(i, j+1).Value = 0
      Else
        Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "NECKUP00__FOZA", "Minimum", 1, -1000)
        Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "NECKUP00__FOZA", "Min_time", 0, 1)
      End If
      i = i + 1
    Else
      Call LogError(Occ_Position & "NECKUP00__FOZA does not exist")
      i = i + 2
    End If
    
    'Neck X moment
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "NECKUP00__MOXB", "Max_moment", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "NECKUP00__MOXB", "Max_time", 0, 1)
    i = i + 1
    
    'Neck Z moment
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "NECKUP00__MOZB", "Max_moment", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "NECKUP00__MOZB", "Max_time", 0, 1) 
    i = i + 1
    
    'Deflection
    temp_index = Data.Root.ChannelGroups(Occ_Position & "_Shoulder").Properties("index").Value
        
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "SHRILE00__DSYB", "Minimum", 0, -1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "SHRILE00__DSYB", "Min_time", 0, 1) 
    i = i + 1
    
    'Y Force
    If Data.Root.ChannelGroups(temp_index).Channels.Exists(Occ_Position & "SHLDLE00__FOYA") Then 
      If Data.Root.ChannelGroups(temp_index).Channels(Occ_Position & "SHLDLE00__FOYA").Properties("Maximum").Value < 0 Then
        ExcelSheet.Cells(i, j).Value = 0
        ExcelSheet.Cells(i, j+1).Value = 0
      Else
        Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "SHLDLE00__FOYA", "Maximum", 1, 1000)
        Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "SHLDLE00__FOYA", "Max_time", 0, 1)        
      End If
    Else
      Call LogError(Occ_Position & "SHLDLE00__FOYA does not exist")
    End If 
    i = i + 1

    'Ribs
    temp_index = Data.Root.ChannelGroups(Occ_Position & "_Torso").Properties("index").Value
    
    Dim Rib_name, Rib_count
    Rib_name = array("TRRI01LE__", "TRRI02LE__", "TRRI03LE__", "ABRI01LE__", "ABRI02LE__")
    
    'Peak rib deflections 
    For Rib_count = 0 to 4
      Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & Rib_name(Rib_count) & "DSYB", "Minimum", 0, -1)
      Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & Rib_name(Rib_count) & "DSYB", "Min_time", 0, 1) 
      i = i + 1
    Next
    
    'Average rib deflection
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannelGroup, temp_index, "", "Avg_rib_defl", 0, -1)
    i = i + 1
    
    'Peak rib deflection rate
    For Rib_count = 0 to 4 
      Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & Rib_name(Rib_count) & "YVF", "Maximum", 2, 1)
      Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & Rib_name(Rib_count) & "YVF", "Max_time", 0, 1) 
      i = i + 1
    Next
    
    'Peak V*C
    For Rib_count = 0 to 4
      Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & Rib_name(Rib_count) & "VC", "Maximum", 2, 1)
      Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & Rib_name(Rib_count) & "VC", "Max_time", 0, 1) 
      i = i + 1
    Next

    'Pelvis & Left Femur
    temp_index = Data.Root.ChannelGroups(Occ_Position & "_Pelvis_Femur").Properties("index").Value
    
    'Pelvis lateral acceleration
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "PELV0000__ACYA", "Maximum", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "PELV0000__ACYA", "Max_time", 0, 1) 
    i = i + 1
    
    'Iliac force
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "ILACLE00__FOYB", "Maximum", 1, 1000)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "ILACLE00__FOYB", "Max_time", 0, 1) 
    i = i + 1
    
    'Acetabulum force
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "ACTBLE00__FOYB", "Maximum", 1, 1000)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "ACTBLE00__FOYB", "Max_time", 0, 1) 
    i = i + 1
    
    'Combined iliac and acetabulum force
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "COMBINED_PELV_FOY", "Maximum", 1, 1000)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "COMBINED_PELV_FOY", "Max_time", 0, 1) 
    i = i + 1
    
    'Femur Y force 3ms clip
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "FEMRLELO__FOYB", "Clip_3_ms", 1, 1000)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "FEMRLELO__FOYB", "Clip_3_ms_T1", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+2, eSearchChannel, temp_index, Occ_Position & "FEMRLELO__FOYB", "Clip_3_ms_T2", 0, 1)
    i = i + 1
    
    'Femur L-M moment 3ms clip
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "FEMRLELO__MOXB", "Clip_3_ms", 1, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "FEMRLELO__MOXB", "Clip_3_ms_T1", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+2, eSearchChannel, temp_index, Occ_Position & "FEMRLELO__MOXB", "Clip_3_ms_T2", 0, 1)
    i = i + 1
  
    'Femur A-P moment 3ms clip
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, Occ_Position & "FEMRLELO__MOYB", "Clip_3_ms", 1, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, Occ_Position & "FEMRLELO__MOYB", "Clip_3_ms_T1", 0, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+2, eSearchChannel, temp_index, Occ_Position & "FEMRLELO__MOYB", "Clip_3_ms_T2", 0, 1)
    i = i + 1
  
    'MDB and VEHICLE
    i = 42  'row
    j = 3  'column
    
    'MDB Delta-V
    dim MDBGroup
    If Data.Root.ChannelGroups.Exists(Data.Root.Properties("Test_ID").Value & "_MDB") Then
      Set MDBGroup = Data.Root.ChannelGroups(Data.Root.Properties("Test_ID").Value & "_MDB")
      Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, MDBGroup.Properties("index").Value, "M0MBARCG0000VAXC", "Delta_V", 1, 1)
      Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, MDBGroup.Properties("index").Value, "M0MBARCG0000VAXC", "Min_time", 0, 1) 
    Else
      Call LogError("No MDB data to export")
    End If 
    i = i + 1         
    
    'Vehicle Delta-V
    temp_index = Data.Root.ChannelGroups("Vehicle").Properties("index").Value 
    
    'Vehicle front accel
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, "10VEHCRIFR__VEY", "Maximum", 1, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, "10VEHCRIFR__VEY", "Max_time", 0, 1) 
    i = i + 1
  
    'Vehicle rear accel
    Call TransferPropToCellRef(ExcelSheet, i, j, eSearchChannel, temp_index, "10VEHCRIRE__VEY", "Maximum", 1, 1)
    Call TransferPropToCellRef(ExcelSheet, i, j+1, eSearchChannel, temp_index, "10VEHCRIRE__VEY", "Max_time", 0, 1) 
    i = i + 1
  End Sub

  'This sub transfers film analysis data
  Sub FilmAnalysisExport(HasDriver, HasPassenger)
    Dim CellName, CustomProp
    Dim ExcelSheet : Set ExcelSheet = objExcel.Workbooks(DiademExcelWorkbook).Sheets("Export From DIAdem")
    ExcelSheet.Select
    Call TransferPropToNamedCell(ExcelSheet, "ImpactVelocity", eSearchFile, "", "", "Impact_velocity", 1, 1)
    
    If HasDriver Then
      Call TransferPropToNamedCell(ExcelSheet, "DriverShoulderLoadingTime", eSearchChannelGroup, data.Root.ChannelGroups("11_Shoulder").Properties("index").Value, "", "Shoulder_loading_time", 1, 1)
      Call TransferPropToNamedCell(ExcelSheet, "DriverTorsoLoadingTime", eSearchChannelGroup, data.Root.ChannelGroups("11_Torso").Properties("index").Value, "", "Torso_loading_time", 1, 1)
      Call TransferPropToNamedCell(ExcelSheet, "DriverPelvisLoadingTime", eSearchChannelGroup, data.Root.ChannelGroups("11_Pelvis_Femur").Properties("index").Value, "", "Pelvis_loading_time", 1, 1)
    End If
    If HasPassenger Then
      Call TransferPropToNamedCell(ExcelSheet, "PassengerShoulderLoadingTime", eSearchChannelGroup, data.Root.ChannelGroups("14_Shoulder").Properties("index").Value, "", "Shoulder_loading_time", 1, 1)
      Call TransferPropToNamedCell(ExcelSheet, "PassengerTorsoLoadingTime", eSearchChannelGroup, data.Root.ChannelGroups("14_Torso").Properties("index").Value, "", "Torso_loading_time", 1, 1)
      Call TransferPropToNamedCell(ExcelSheet, "PassengerPelvisLoadingTime", eSearchChannelGroup, data.Root.ChannelGroups("14_Pelvis_Femur").Properties("index").Value, "", "Pelvis_loading_time", 1, 1)
    End If
    
    If Data.Root.ChannelGroups.Exists("Film_Analysis") = False Then
      Exit Sub
    End If
    
    For Each CellName In objExcel.Workbooks(DiademExcelWorkbook).Names
      For Each CustomProp In Data.Root.ChannelGroups("Film_Analysis").Properties
        If CellName.Name = CustomProp.Name Then
          ExcelSheet.Range(CellName.Name).Value = CustomProp.Value
        End If
      Next
    Next
  End Sub

  'Run Excel TDMS Add-in
  Sub Run_Excel_TDMS_Addin()
  
    Dim TdmAddIn, TdmsFilePath, Template_Workbook, TDMS_Workbook, TDMS_filename
    
    TDMS_filename = ".TDMS"
    
    'objExcel = CreateObject("Excel.Application")
    Template_Workbook = objExcel.ActiveWorkbook.Name
    
    TdmsFilePath = "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\Side-Impact Crash Tests\"
    
    
    On Error Resume Next
    
    objExcel.Visible = True
    If Err.number <> 0 Then Call AutoQuit("Could not contact Excel on this computer")
    Set TdmAddIn = objExcel.COMAddIns.Item("ExcelTDM.TdmAddin")
    If Err.number <> 0 Then Call AutoQuit("Could not access the TDM Excel Addin" & vbCrLf & vbCrLf & "(TDM Excel Addin version 3.1 or later is required)")
    objExcel.Visible = True
    Call TdmAddIn.Object.ImportFile(TdmsFilePath & TESTID & "\DATA\DIAdem\" & TESTID & TDMS_filename, True)
    If Err.number <> 0 Then Call AutoQuit("Could not load the desired data file" & vbCrLf & vbCrLf & """" & TdmsFilePath & """")
    Call objExcel.ActiveWindow.Activate
    On Error GoTo 0
 
  End Sub