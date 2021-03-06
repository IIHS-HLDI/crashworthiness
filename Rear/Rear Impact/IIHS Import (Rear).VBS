'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/23/2010 13:28:34
'-- Author: 
'-- Comment: Rear Impact Import
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

  'Clear data
  Call Data.Root.Clear()
  
  'Import Data from TDAS directory
  Dim TDAS_Folder
  TDAS_Folder = "N:\VRC\Shared\Resources\Dummy Information\TDAS\REAR\DATA\"
  Call DataFileLoad(TDAS_Folder & T1 & "\" & T1 & "_DIAdemHeader_BINARY.dat","DD8x","") '... DataFilename,FileImportFilter,ImportAction 
  Data.Root.ChannelGroups(1).Name = "Raw Data"
  
  'Rename Root Directory and Add properties for test ID, vehicle name, and test date
  Data.Root.Name = T1 & " - " & T2
  Call Data.Root.Properties.Add("Test_ID", T1, DataTypeString)
  Call Data.Root.Properties.Add("Vehicle_Name", T2, DataTypeString)
  Call Data.Root.Properties.Add("Test_Date", T3, DataTypeString)
  Call Data.Root.Properties("Description").Value(T4)
  
  
  'Call script to change numeric channels to waveforms, remove offset, and create filtered and caluculated channel groups
  Scriptinclude(CurrentScriptPath & "Change to Waveform and Remove offset (Rear).vbs")
  
  'Call script to filter data
  Scriptinclude(CurrentScriptPath & "Filter Data (Rear).vbs")
  
  'Call script to shift data so the peak sled acceleration occurs at 27 ms.
  Scriptinclude(CurrentScriptPath & "Data shift (Rear).vbs")
  
  'Call script to calculate injury criteria
  Scriptinclude(CurrentScriptPath & "Calculate Injury Criteria (Rear).vbs")
  
  'Call script to calculate delta V for the sled accelerometers
  Scriptinclude(CurrentScriptPath & "Delta V calculations (Rear).vbs")
  
  'Call script to calculate HCT
  Scriptinclude(CurrentScriptPath & "HCT (Rear).vbs")
  
  'Call script to load IIWPG ideal pulse verify test pulse
  Scriptinclude(CurrentScriptPath & "Pulse Specifications (Rear).vbs")
  
  'Call script to determine the max, min, and time values for export to excel
  Scriptinclude(CurrentScriptPath & "Max, Min, Time Values.vbs")
  
  'Call script to calculate dynamic rating
  Scriptinclude(CurrentScriptPath & "Ratings (Rear).vbs")
  
  'Call script that inputs sled test information
  Scriptinclude(CurrentScriptPath & "Sled Test Info.vbs")
  
  'Collapse groups in data portal
  Call UIAutoRefreshSet(True) 
  call Portal.Structure.Collapse(Data.Root)
  
  'Call script to load the view and report
  Scriptinclude(CurrentScriptPath & "Load View and Report (Rear).vbs")

'Copy DAS folder to Side CE directory
  dim filesys
  dim SourceFolder
  dim DestinationFolder
  
  SourceFolder = "N:\VRC\Shared\Resources\Dummy Information\TDAS\Rear\DATA\"
  DestinationFolder = "N:\VRC\Shared\Crashworthiness Program\Rear Impact Sled Tests\"
  
  Set filesys=CreateObject("Scripting.FileSystemObject")
  If filesys.FolderExists(SourceFolder & T1) Then
    filesys.CopyFolder SourceFolder & T1, DestinationFolder & T1 & "\TDAS", 1 
  Else
    msgbox(SourceFolder & T1 & " does not exist.")
  End If


  'Save TDMS file in test directory
  Dim Folder
  Dim Name
  Folder =  "N:\VRC\Shared\Crashworthiness Program\Rear Impact Sled Tests\"
  Name = Data.Root.Properties("name").Value
  Call DataFileSave(Folder & T1 & "\" & "DATA\DIAdem\" & Name & ".tdms","TDMS") '... DataFilename,FileExportFilter 
  
  'Save View file in test directory
  Call View.SaveLayout(Folder & T1 & "\" & "DATA\DIAdem\" & Name & ".TDV")
  
  'Save Report file in test directory
  'Call PicSaveAs(Folder & T1 & "\" & "DATA\DIAdem\" & Name & ".TDR","") '... PicFile,PICFileCode 
  
  'Create PDF summary file
  'Call PicPDFExport(Folder & T1 & "\" & "DATA\DIAdem\" & Name & ".pdf",0) '... PDFFilename,PDFAppend 
  
  'Call script to create excel summary file
  Scriptinclude(CurrentScriptPath & "Transfer Rear Data to Excel.vbs")
  
  'Call script to add pulse data to summary file
  Scriptinclude(CurrentScriptPath & "Transfer Pulse Data to Excel.vbs")
  
  Call WndShow("Report", "Maximize")



