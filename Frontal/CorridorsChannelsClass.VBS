'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/23/2017 13:45:05
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\CorridorClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")

Class CorridorsChannels
  Private iGroupName, oGroup, iChannelDictionary, ioccupantLocation, oConfigFileFactory, iGroupTagName
  Public Property Get ChannelDictionary
    Set ChannelDictionary = iChannelDictionary
  End Property

  Public Default Function Init(occupantLocation, configFileFactory)
    iGroupTagName = "corridorsGroupName"
    iOccupantLocation = occupantLocation
    Set oConfigFileFactory = configFileFactory
    Set iChannelDictionary = CreateObject("Scripting.Dictionary")
    iGroupName = oConfigFileFactory.GetConfigValue("testSettings", iGroupTagName, "name")
    Set oGroup = CreateCorridorsGroup(iOccupantLocation & iGroupName)
    Call LoadChannels
    Set Init = Me
  End Function
 
  Private Sub LoadChannels
    Dim slugs : Set slugs = oConfigFileFactory.GetCorridorList(iGroupTagName)
    Dim slug, key
    For Each slug in slugs
      key = oConfigFileFactory.GetConfigValue("CorridorsMetaData", slug, "key")
      Call iChannelDictionary.Add(key, (New Corridor)(iOccupantLocation, slug, oConfigFileFactory))
    Next
  End Sub
  
  Public Sub FilterChannels
  End Sub
  
  Private Function CreateCorridorsGroup(name)
    If Data.Root.ChannelGroups.Exists(name) Then
      Set CreateCorridorsGroup = Data.Root.ChannelGroups(name)
      Exit Function
    Else
      On Error Resume Next
      Dim path : path = oConfigFileFactory.GetConfigValue("testSettings", iGroupTagName, "path")
      Dim grp : Set grp = DataFileLoad(APPPATH & path, "","")(1) 'DataFileLoad returns a collection of groups
      grp.Name = name
      Dim chn
      For Each chn In grp.Channels
        chn.Name = iOccupantLocation & chn.Name
      Next
      Set CreateCorridorsGroup = grp
      If Err.number <> 0 Then
        LogError("Unable to load corridors: " & Replace(Err.Description, vbCrLf, ""))
      End If
      On Error GoTo 0
      Set grp = Nothing
    End If
  End Function
  
  Private Sub CreateLoadDuration(oLD, corridorTimeChnName, lowerBoundChnName)
    Dim posResultName : posResultName = "PosPulse"
    Dim negResultName : negResultName = "NegPulse"
    Dim ldResultName : ldResultName = "LoadDuration"
    Dim ampResultName : ampResultName = "Amplitude"
    
    'Split channel at 0
    Call ChnSplitAtValue(oLD.SourceChannel,0,0,"/" & posResultName,"/" & negResultName) 
    Call ChnLinScale("/" & negResultName ,"/" & negResultName,-1,0)  
    
    Dim timeChn
    If oLD.InvertSource Then 
      timeChn = WfChnToChn("/" & negResultName, true, "WfXRelative")(0)
      Call ChnDurPosCalc(timeChn,"/" & negResultName, "in one peak", "/" & ldResultName, "/" & ampResultName) 
    Else
      timeChn = WfChnToChn("/" & negResultName, true, "WfXRelative")(0)
      Call ChnDurPosCalc(timeChn,"/" & posResultName, "in one peak", "/" & ldResultName,"/" & ampResultName) 
    End If

    Call ChnMapLinCalc("/" & ldResultName, "/" & ampResultName, "/" & corridorTimeChnName, "/" & oLD.CorridorChannelName, 1, "const. value", NOVALUE, "analogue") 
    Call ChnToWfChn("/" & corridorTimeChnName,"/" & oLD.CorridorChannelName, 0) 
    Call ChnNovHandle("/" & corridorTimeChnName, "/" & oLD.CorridorChannelName, "SetValue", "XY", 1, 0, 0)
    Call ChnPulseLimit("/" & corridorTimeChnName, "/" & oLD.CorridorChannelName, "/" & corridorTimeChnName, "/" & oLD.ReferenceChannel.Name, "/" & corridorTimeChnName, "/" & lowerBoundChnName) 
    
    Dim resultText
    If pulselimit Then
      resultText = "PASS"
    Else
      resultText = "FAIL"
      Call oLD.CorridorChannel.Properties.Add("FAIL_POINT", PulseLimitTIdx,DataTypeInt32)
    End If
    
    Call oLD.CorridorChannel.Properties.Add("Corridor_Result", resultText,DataTypeString)
    
    Call oGroup.Channels.Remove(Data.GetChannel(timeChn).Name)
    Call oGroup.Channels.Remove(posResultName)
    Call oGroup.Channels.Remove(negResultName)
    Call oGroup.Channels.Remove(ldResultName)
    Call oGroup.Channels.Remove(ampResultName)
    
  End Sub
    
  Public Sub CalculateInjuryCriteria
    On Error Resume Next
    oGroup.Activate
    Dim timeChnName : timeChnName = iChannelDictionary.Item("time").CorridorChannelName
    Dim lowerBoundChnName : lowerBoundChnName = iChannelDictionary.Item("lowerBound").CorridorChannelName
    Dim item
    For Each item In iChannelDictionary.Items
      If item.IsCalculated Then 
        Call CreateLoadDuration(item, timeChnName, lowerBoundChnName)
        If Not Err.number = 0 Then 
          LogError("Calculation of " & item.Description & " corridor failed. Message: " & Replace(Err.Description, vbCrLf, ""))
          Err.Clear
        End If
      End If
    Next 
    On Error Goto 0
  End Sub
  
  Public Sub ExportProperties(sheet)
    sheet.Select
    
    Dim exportCellName : exportCellName = "CorridorOutput"
    Dim i : i = 0
    Dim resultString, item
    
    For Each item In iChannelDictionary.Items
      If item.IsCalculated Then
        If item.CorridorChannel.Properties("Corridor_Result").Value = "FAIL" Then
          If i > 0 Then
            resultString = resultString & ", " & item.Description
          Else
            resultString = item.Description
          End If
          i = i + 1
        End If
      End If
    Next 
    
    If i = 0 Then
      sheet.Range(exportCellName).Value = "No load-duration corridors were exceeded."
    ElseIf i = 1 Then
      sheet.Range(exportCellName).Value = "Load duration corridor exceeded for the following channel: " & resultString
    Else
      sheet.Range(exportCellName).Value = "Load duration corridor exceeded for the following channels: " & resultString
    End If
  End Sub
End Class