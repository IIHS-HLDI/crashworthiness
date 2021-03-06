'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/07/2010 13:34:32
'-- Author: 
'-- Comment: Front Crash Processing
'-------------------------------------------------------------------------------
Option Explicit
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\Profiles\SmallOverlapTestProfileClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\Profiles\ModerateOverlapTestProfileClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\Profiles\RightSideSmallOverlapTestProfileClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\Profiles\CustomTestProfileClass.VBS")

Call DataFileLoad(DataReadPath  & T1 & "_DIAdemHeader_BINARY.dat","DD8x","") '... DataFilename,FileImportFilter,ImportAction 

Call Data.Root.Properties.Add("Test ID", T1, DataTypeString)
Call Data.Root.Properties.Add("Vehicle_Name", T2, DataTypeString)
Call Data.Root.Properties.Add("Test_Date", T3, DataTypeString)
Call Data.Root.Properties("Description").Value(T4)
Call Data.Root.Properties.Add("Impact_velocity",round(T5,1),DataTypeFloat64)
Data.Root.Properties("author").Value = "IIHS"
Call Data.Root.Properties.Add("Dummy_Type", "H350", DataTypeString)

Call CleanOldFrontalDataScript
Call PreFilterRawData(Data.Root.ChannelGroups(1), T1, T4)
  Dim testTypeId
  If B5 Then
    testTypeId = 999
  Else
    testTypeId = GetTestTypeId(T1)
  End If 
Dim profile

Call Data.Root.Properties.Add("HasDriver", B3)
Call Data.Root.Properties.Add("HasPassenger", B4)
  
Select Case testTypeId
  Case 1
    Set profile = (New ModerateOverlapTestProfile)(T1)
  Case 6
    Set profile = (New SmallOverlapTestProfile)(T1)
  Case 11
    Set profile = (New RightSideSmallOverlapTestProfile)(T1)
  Case 999
    Set profile = (New CustomTestProfile)(T1)
End Select
  
profile.Run
LaunchLog

Call UIAutoRefreshSet(True) 
Call Portal.Structure.Collapse(Data.Root) 
Call UIAutoRefreshSet(False) 