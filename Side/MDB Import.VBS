'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/16/2013 14:35:37
'-- Author: ---
'-- Comment: ---
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

  'Copy MDB Slice folder to Side CE directory
  dim filesys
  dim SourceFolder
  dim DestinationFolder
  
  SourceFolder = "C:\DTS\SLICEWare\1.08.0816\DATA\"
  DestinationFolder = "N:\VRC\Shared\Crashworthiness Program\Side-Impact Crash Tests\"
  
  Set filesys=CreateObject("Scripting.FileSystemObject")
  If filesys.FolderExists(SourceFolder & T1 & " MDB") Then
    filesys.CopyFolder SourceFolder & T1 & " MDB", DestinationFolder & T1 & "\DATA\DAS", 1 
    Call LoadProcessMDB
  Else
    Call Data.Root.ChannelGroups.Add("MDBNotPresent")
    msgbox(SourceFolder & T1 & " MDB" & "does not exist.")
  End If

  Sub LoadProcessMDB()
    'Load MDB file from test directory
    Call DataFileLoad(DestinationFolder & T1 & "\DATA\DAS\" & T1 & " MDB.dat","DD8x","Load") '... DataFilename,FileImportFilter,ImportAction 
    
    'Convert channels to waveform, chop channels and remove offset
    Call ChnToWfChn("/Time axis","'/M0MBARCG0000ACZ0' - '/M0MBARCG0000ACX0'",0) '... X,ChnNoStr,XChnDelete
    
    Call ChopChannels(Data.Root.ChannelGroups(2).Channels)
    Call Data.Root.ChannelGroups(2).Channels.Remove("Time axis")
    Call ChnLinGenImp("Time axis",6000,-0.1,0.0001,"")
    Call Data.Root.ChannelGroups(2).Channels("Time axis").Properties.Add("Test ID", T1, DataTypeString)
    Call Data.Root.ChannelGroups(2).Channels("Time axis").Properties.Add("Test Type", T4, DataTypeString) 
    Call RemoveOffset(Data.Root.ChannelGroups(2).Channels, 100)
    
    'Change unit to Diadem's acceleration gn
    Data.Root.ChannelGroups(2).Channels("M0MBARCG0000ACZ0").Properties("unit_string").Value = "gn"
    Data.Root.ChannelGroups(2).Channels("M0MBARCG0000ACY0").Properties("unit_string").Value = "gn"
    Data.Root.ChannelGroups(2).Channels("M0MBARCG0000ACX0").Properties("unit_string").Value = "gn"
    
    'Create filtered channels
    Call ChnCFCFiltCalc("","/M0MBARCG0000ACX0","/M0MBARCG0000ACXD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
    Call ChnCFCFiltCalc("","/M0MBARCG0000ACY0","/M0MBARCG0000ACYD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
    Call ChnCFCFiltCalc("","/M0MBARCG0000ACZ0","/M0MBARCG0000ACZD","CFC_60",0,"EndPoints",10) '... XW,Y,E,CFCFiltType,Fir100RemoveBias,CFCPreEventType,CFCFreeValue 
    Call ChnCFCFiltCalc("","/M0MBARCG0000ACX0","/M0MBARCG0000ACXC","CFC_180",0,"EndPoints",10)
    
    'Calculate DeltaV
    'Extract pre t=0 data from MDB_ACXC (data filtered at CFC180)
    Dim i, temp_value, channel_target_point, channel_target_point2, integration_endtime
    Dim channel_length, channel_offset, channel_step
    
    channel_length = Data.Root.ChannelGroups(2).Channels("M0MBARCG0000ACXC").Properties("length").Value
    channel_offset = Data.Root.ChannelGroups(2).Channels("M0MBARCG0000ACXC").Properties("wf_start_offset").Value
    channel_step = Data.Root.ChannelGroups(2).Channels("M0MBARCG0000ACXC").Properties("wf_increment").Value
    
    'Find min value point of vehicle acceleration curve, then determine point that is 1000 points beyond the zero crossing
    For i = 1 to (channel_length - 1)
        temp_value = Data.Root.ChannelGroups(2).Channels("M0MBARCG0000ACXC").Values(i)
        If temp_value = Data.Root.ChannelGroups(2).Channels("M0MBARCG0000ACXC").Properties("minimum").Value THEN
          channel_target_point = i 
          Exit For
        End If
    Next
    
    For i = channel_target_point to (channel_length - 1)
        temp_value = Data.Root.ChannelGroups(2).Channels("M0MBARCG0000ACXC").Values(i)
        If temp_value >= 0 THEN
          channel_target_point2 = i + 999
          Exit For
        End If
    Next
      
    integration_endtime = channel_offset + (channel_step * channel_target_point2)
        
    Call ChnTimeAreaCopy("","/M0MBARCG0000ACXC","","/TimeRangeCopyY",0,integration_endtime) '... XW,Y,E,E,TimeAreaCopyBeg,TimeAreaCopyEnd  
    
    'Integrate data
    Call ChnIntegrate("","/TimeRangeCopyY","/Integrated") '... XW,Y,E 
    'Scale integrated data by 35.316 (converts G-s to km/h)
    Call Calculate("Ch(""/M0MBARCG0000VAXC"")=Ch(""/Integrated"")*9.81*3.6+T5",NULL,NULL,"") '... CalculateFormula,CalculateSymbols,CalculateValues,CalculateTargetUnit
    Data.Root.ChannelGroups(2).Channels("M0MBARCG0000VAXC").Properties("wf_xname").Value = "Time_MDB"
    Data.Root.ChannelGroups(2).Channels("M0MBARCG0000VAXC").Properties("unit_string").Value = "km/h"
    
    Dim deltaV
    deltaV = Data.Root.ChannelGroups(2).Channels("M0MBARCG0000VAXC").Properties("minimum").Value - Data.Root.ChannelGroups(2).Channels("M0MBARCG0000VAXC").Properties("maximum").Value
    
    Call Data.Root.ChannelGroups(2).Channels("M0MBARCG0000VAXC").Properties.Add("Delta_V", round(deltaV,1), DataTypeFloat64)  
    Call ChnTMinTMaxCalc("","M0MBARCG0000VAXC")
    Call Data.Root.ChannelGroups(2).Channels("M0MBARCG0000VAXC").Properties.Add("Min_time", round(mintime*1000,1), DataTypeFloat64) 

    'Delete channels that aren't needed
    Call Data.Root.ChannelGroups(2).Channels.Remove("TimeRangeCopyY")
    Call Data.Root.ChannelGroups(2).Channels.Remove("Integrated")
    
    Data.Root.ChannelGroups(2).Name = T1 & "_MDB"
  End Sub

  Sub RemoveOffset(channels, numberOfPoints)
      dim chn
      For Each chn In channels
          If Chn.Name <> "Time axis" Then
              Call ChnOffset(chn, chn, numberOfPoints, "mean value offset")
          End If
      Next
  End Sub
  Sub ChopChannels(channels)
      dim chn
      For Each chn In channels
          If Chn.Name <> "Time axis" Then
              Call ChnTimeAreaCopy("", Chn, "", Chn, -0.1, 0.5)
          End If
      Next
  End Sub
