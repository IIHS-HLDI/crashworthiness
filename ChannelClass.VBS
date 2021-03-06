'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 08/02/2016 17:55:47
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")

Class Channel
  Private iIsoCode, iFilter, iOccupantLocation, iChannelSlug, iRawGroupName, iPreFilteredChannelName, iFilteredChannelName, iIsCalculated, oConfigFileFactory, iGroupTag, oGroup, iExportList
  
  Public Property Get FilteredChannel
    If oGroup.Channels.Exists(FilteredChannelName) Then
        Set FilteredChannel = oGroup.Channels(FilteredChannelName)
    Else
        Set FilteredChannel = Nothing
    End If
  End Property 

  Public Property Get PreFilteredChannel
    If cbool(iIsCalculated) Then 
      Set PreFilteredChannel = oGroup.Channels(iPreFilteredChannelName)
    Else
      If Data.Root.ChannelGroups(iRawGroupName).Channels.Exists(iPreFilteredChannelName) Then
          Set PreFilteredChannel = Data.Root.ChannelGroups(iRawGroupName).Channels(iPreFilteredChannelName)
      Else
          Set PreFilteredChannel = Nothing
      End If
    End If
  End Property 

  Public Property Get PreFilteredChannelName
    PreFilteredChannelName = iPreFilteredChannelName
  End Property 

  Public Property Get FilteredChannelName
    If Not IsEmpty(iFilteredChannelName) Then
      FilteredChannelName = iOccupantLocation & iFilteredChannelName
    Else
      FilteredChannelName = iFilteredChannelName
    End If
  End Property
  
  Public Property Get ExportList
    Set ExportList = iExportList
  End Property
  
  Public Default Function Init(occupantLocation, channelSlug, configFileFactory)
    iOccupantLocation = occupantLocation
    iChannelSlug = channelSlug
    Set oConfigFileFactory = configFileFactory
    Call GetChannelInfo
    If iGroupTag = "vehicleGroupName" Then 
      Set oGroup = Data.Root.ChannelGroups(Left(iOccupantLocation, 1) & oConfigFileFactory.GetConfigValue("testSettings", iGroupTag, "name"))
    Else
      Set oGroup = Data.Root.ChannelGroups(iOccupantLocation & oConfigFileFactory.GetConfigValue("testSettings", iGroupTag, "name"))
    End If
    Set Init = Me
  End Function
  
  Private Sub GetChannelInfo
    iIsoCode = oConfigFileFactory.GetConfigValue("ChannelMetaData", iChannelSlug, "isoCode")
    iFilter = Ucase(oConfigFileFactory.GetConfigValue("ChannelMetaData", iChannelSlug, "filter"))
    iRawGroupName = GetRawGroupName(oConfigFileFactory.GetConfigValue("testSettings", "rawGroupName", "name"))
    iPreFilteredChannelName = iOccupantLocation & iIsoCode
    iFilteredChannelName = oConfigFileFactory.GetConfigValue("ChannelMetaData", iChannelSlug, "filteredName")
    iIsCalculated = oConfigFileFactory.GetConfigValue("ChannelMetaData", iChannelSlug, "isCalculated")
    iGroupTag = oConfigFileFactory.GetConfigValue("ChannelMetaData", iChannelSlug, "groupTag")
    Set iExportList = oConfigFileFactory.GetExportList(iChannelSlug)
  End Sub
  
  Public Sub FilterChannel
    dim filterString
    
    If iIsCalculated Then Exit Sub
    
    If IsEmpty(iFilter) Then
      LogError("No filter class loaded")
      Exit Sub
    End If
    
    Select Case iFilter
      Case "A"
        filterString = "CFC_1000"
      Case "B"
        filterString = "CFC_600"
      Case "C"
        filterString = "CFC_180"
      Case "D"
        filterString = "CFC_60"
    End Select
    
    On Error Resume Next
      Call ChnCFCFiltCalc("", iRawGroupName & "/" & iPreFilteredChannelName, "/" & iOccupantLocation & iFilteredChannelName, filterString, 0, "EndPoints", 10) 
      If Err.number <> 0 Then Call LogError("Unable to filter channel " & iPreFilteredChannelName  & ". Message: " & Replace(Err.Description, vbCrLf, ""))
    On Error Goto 0
    
  End Sub
  
  Public Sub ExportProperties(ExcelSheet)
    ExcelSheet.Select
    Dim item, exportData, chn
    For Each exportData In ExportList
      If exportData.UseFilteredChannel Then
        Set chn = FilteredChannel
      Else
        Set chn = PreFilteredChannel
      End If
      Call TransferPropToExcel(ExcelSheet, exportData.Cell , chn, exportData.PropertyName, exportData.RoundValue, exportData.Divisor)
    Next
    
    Set chn = Nothing
  End Sub
End Class