'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/07/2010 13:34:32
'-- Author: 
'-- Comment: Side Crash Processing
'-------------------------------------------------------------------------------
  Dim SideCE_path : SideCE_path = "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\Side-Impact Crash Tests\"
  Dim Occ_Position
  View.AutoRefresh = TRUE
  
  Call DataFileLoad(DataReadPath & "\" & T1 & "_DIAdemHeader_BINARY.dat","DD8x","") 'THIS IS THE CORRECTED USE SO THAT DIADEM 2014 WILL PROCESS DATA
  
  Data.Root.Name = T1
  Data.Root.ChannelGroups(1).Name = T1 & "_Raw_Data"
  Call Data.Root.Properties.Add("Test ID", T1, DataTypeString)
  Call Data.Root.Properties.Add("Vehicle_Name", T2, DataTypeString)
  Call Data.Root.Properties.Add("Test_Date", T3, DataTypeString)
  Call Data.Root.Properties("Description").Value(T4)
  Call Data.Root.Properties.Add("Impact_velocity",round(T5,1),DataTypeFloat64)
  Data.Root.Properties("author").Value = "IIHS"
  Call Data.Root.Properties.Add("Data_Path", SideCE_path & T1, DataTypeString)
  
  'Call Script to convert data to waveforms and perform offset removal
  ScriptInclude(CurrentScriptPath & "Side Change to Waveform and Remove offset")
  'Import and Filter MDB Data
  ScriptInclude(CurrentScriptPath & "MDB import")
  'Filter all data
  ScriptInclude(CurrentScriptPath & "Filter SID2S Data")
  'Calcualte injuries
  ScriptInclude(CurrentScriptPath & "Calculate SID2S injury criteria")
  '----------------------------------------
  'SAVE TDMS File to network
  
  Call DataFileSave(SideCE_path & T1 & "\DATA\DIAdem\" & T1 & ".TDMS","TDMS") '... DataFilename,FileExportFilter 
  '----------------------------------------
  'Copy layout view from network and open
  Call Copy_View_Layout
  '----------------------------------------
  'Transfer summary data to Excel 
  ScriptInclude(CurrentScriptPath & "Transfer Side CE to Excel")
  '-----------------------------------------
  
  Call Copy_DAS_Folder
  
  'Collapse groups in data portal
  Call UIAutoRefreshSet(True) 
  Call Portal.Structure.Collapse(Data.Root) 
  Call UIAutoRefreshSet(False) 
  
  '----------------------------------------
  '----------------------------------------
  'Subroutines shown below
  '----------------------------------------
  '----------------------------------------
  'Copy Side View layout from network to test folder
  Sub Copy_View_Layout
    Dim SourceFolder
    SourceFolder = CurrentScriptPath
    Set filesys=CreateObject("Scripting.FileSystemObject")
    
    IF B1 = TRUE THEN  'Copy driver unfiltered layout file
      If filesys.FileExists(SourceFolder & "SIDE_UNFILT_DR.TDV") Then
        filesys.CopyFile SourceFolder & "SIDE_UNFILT_DR.TDV", SIDECE_path & T1 & "\DATA\DIAdem\SIDE_UNFILT_DR.TDV"
        Call View.LoadLayout( SIDECE_path & T1 & "\DATA\DIAdem\" & "SIDE_UNFILT_DR.TDV")
      Else
        msgbox(SourceFolder & "SIDE_UNFILT_DR.TDV" & " does not exist.")
      End If
    END IF
    
    IF B2 = TRUE THEN  'Copy rear passenger unfiltered layout file
      If filesys.FileExists(SourceFolder & "SIDE_UNFILT_REAR.TDV") Then
        filesys.CopyFile SourceFolder & "SIDE_UNFILT_REAR.TDV", SIDECE_path & T1 & "\DATA\DIAdem\SIDE_UNFILT_REAR.TDV"
        Call View.LoadLayout( SIDECE_path & T1 & "\DATA\DIAdem\" & "SIDE_UNFILT_REAR.TDV")
      Else
        msgbox(SourceFolder & "SIDE_UNFILT_REAR.TDV" & " does not exist.")
      End If
    END IF
    
    'Copy main layout file
    If filesys.FileExists(SourceFolder & "SIDE_CE.TDV") Then
      filesys.CopyFile SourceFolder & "SIDE_CE.TDV", SIDECE_path & T1 & "\DATA\DIAdem\" & T1 & ".TDV"
      Call View.LoadLayout( SIDECE_path & T1 & "\DATA\DIAdem\" & T1 & ".TDV")
    Else
      msgbox(SourceFolder & "SIDE_CE.TDV" & " does not exist.")
    End If
  
  End Sub
  '----------------------------------------
  'Copy DAS folder to Side CE directory
  Sub Copy_DAS_Folder
  
    dim filesys
    dim SourceFolder
    dim DestinationFolder
    
    SourceFolder = "\\IIHS\IIHSDrive\VRC\Shared\Resources\Dummy Information\TDAS\SIDE\DATA\"
    DestinationFolder = "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\Side-Impact Crash Tests\"
    
    Set filesys=CreateObject("Scripting.FileSystemObject")
    If filesys.FolderExists(SourceFolder & T1) Then
      filesys.CopyFolder SourceFolder & T1, DestinationFolder & T1 & "\DATA\DAS", 1 
    Else
      msgbox(SourceFolder & T1& " does not exist.")
    End If
  End Sub
'----------------------------------------




