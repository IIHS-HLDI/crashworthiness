'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/23/2017 10:14:45
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ChannelClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")

Class ChestChannels
  Private iGroupName, oGroup, iChannelDictionary, ioccupantLocation, oConfigFileFactory, iGroupTagName
  Public Property Get ChannelDictionary
    Set ChannelDictionary = iChannelDictionary
  End Property

  Public Default Function Init(occupantLocation, configFileFactory)
    iGroupTagName = "chestGroupName"
    iOccupantLocation = occupantLocation
    Set oConfigFileFactory = configFileFactory
    Set iChannelDictionary = CreateObject("Scripting.Dictionary")
    iGroupName = oConfigFileFactory.GetConfigValue("testSettings", iGroupTagName, "name")
    Set oGroup = CreateGroup(iOccupantLocation & iGroupName)
    Call LoadChannels
    Set Init = Me
  End Function
 
  Private Sub LoadChannels
    Dim slugs : Set slugs = oConfigFileFactory.GetChannelList(iGroupTagName)
    Dim slug, key
    For Each slug in slugs
      key = oConfigFileFactory.GetConfigValue("ChannelMetaData", slug, "key")
      Call iChannelDictionary.Add(key, (New Channel)(iOccupantLocation, slug, oConfigFileFactory))
    Next
  End Sub
  
  Public Sub FilterChannels
    Dim item 
    oGroup.Activate
    For Each item in iChannelDictionary.Items
      item.FilterChannel
    Next
  End Sub
  
  Public Sub CalculateInjuryCriteria
    On Error Resume Next
    oGroup.Activate
    
    'Chest resultant acceleration
    Dim chestX : Set chestX = iChannelDictionary.Item("chestX").FilteredChannel
    Dim chestY : Set chestY = iChannelDictionary.Item("chestY").FilteredChannel
    Dim chestZ : Set chestZ = iChannelDictionary.Item("chestZ").FilteredChannel
        
    Dim chestR
    Dim chestRName : chestRName = iChannelDictionary.Item("chestR").PreFilteredChannelName
    If oGroup.Channels.Exists(chestRName) Then
      Set chestR = oGroup.Channels(chestRName)
    Else
      Set chestR = oGroup.Channels.Add(chestRName, DataTypeChnFloat64)  
    End If
    
    If Not Err.number = 0 Then 
      LogError("Calculation of chest resultant failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    'Resultant Acceleration 3ms clip
    Call ChnXYZAbsValue(chestX, chestY, chestZ, chestR)
    Call ChnXMSCalc("",chestR,3,0,"in one peak",0)
    
    Call chestR.Properties.Add("Test ID", T1, DataTypeString)
    Call chestR.Properties.Add("Test Type", T4, DataTypeString)
    Call chestR.Properties.Add("Clip_3_ms", XmsAcceleration, DataTypeFloat64)
    Call chestR.Properties.Add("Clip_3_ms_T1", XmsTimeBegin, DataTypeFloat64)
    Call chestR.Properties.Add("Clip_3_ms_T2",XmsTimeEnd , DataTypeFloat64)
    
    Set chestX = Nothing
    Set chestY = Nothing
    Set chestZ = Nothing
    Set chestR = Nothing
    
    If Not Err.number = 0 Then 
      LogError("Calculation of chest 3ms clip failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    'Peak deflection time
    Dim oChestDeflection : Set oChestDeflection = iChannelDictionary.Item("chestDeflection")
    Call ChnTMinTMaxCalc("","/" & oChestDeflection.FilteredChannelName) 
    Call oChestDeflection.FilteredChannel.Properties.Add("Min_time", mintime*1000, DataTypeFloat64)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of chest deflection failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    'Viscous criterion
    Dim chestVCName : chestVCName = iChannelDictionary.Item("chestVC").PreFilteredChannelName
    Call ChnVCCalc("", oChestDeflection.PrefilteredChannel, "/" & chestVCName, 229, 1.3, "SAE_J1727") 
    Call ChnUnitConvert("/" & chestVCName,"/" & chestVCName, "m/s") 
    Call ChnTMinTMaxCalc("", "/" & chestVCName) 
    Call oGroup.Channels(chestVCName).Properties.Add("Test ID", T1, DataTypeString)
    Call oGroup.Channels(chestVCName).Properties.Add("Test Type", T4, DataTypeString)
    Call oGroup.Channels(chestVCName).Properties.Add("Max_time", maxtime*1000, DataTypeFloat64)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of chest vc failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
        
    'Deflection Rate
    Dim chestDeflectionFilterDName : chestDeflectionFilterDName = "ChestDeflectionFilterD"
    Call ChnCFCFiltCalc("", oChestDeflection.PrefilteredChannel, "/" & chestDeflectionFilterDName,"CFC_60",0,"EndPoints",10)
    
    Dim deflectionRateName : deflectionRateName = iChannelDictionary.Item("chestDeflectionRate").PreFilteredChannelName
    Call ChnDifferentiate("","/" & chestDeflectionFilterDName,"","/" & deflectionRateName)
    Call ChnUnitConvert("/" & deflectionRateName,"/" & deflectionRateName,"m/s") 
    Call ChnTMinTMaxCalc("","/" & deflectionRateName) 
    Call oGroup.Channels(deflectionRateName).Properties.Add("Test ID", T1, DataTypeString)
    Call oGroup.Channels(deflectionRateName).Properties.Add("Test Type", T4, DataTypeString)
    Call oGroup.Channels(deflectionRateName).Properties.Add("Min_time", mintime*1000, DataTypeFloat64)
    oGroup.Channels(deflectionRateName).Properties("wf_xname").Value = "Time_CHST0000__VEXD"
    Call oGroup.Channels.Remove(chestDeflectionFilterDName)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of chest deflection rate failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
   
    On Error Goto 0 
    Set oChestDeflection = Nothing
    
  End Sub
  
  Public Sub ExportProperties(sheet)
    sheet.Select
    Dim item, exportData, chn
    For Each item In iChannelDictionary.Items
      item.ExportProperties(sheet)
    Next
    
    Set chn = Nothing
  End Sub
End Class