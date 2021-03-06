'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 08/02/2016 17:55:47
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
    Class ConfigFileFactory
      Private NodeName, CorridorNodeName, ConfigPath, DialogNodeName

      Private Sub Class_Initialize()
        NodeName = "ChannelMetaData"
        CorridorNodeName = "CorridorsMetaData"
        DialogNodeName = "DialogMetaData"
      End Sub
      
      Public Default Function Init(testTypeId)
        Select Case testTypeid
          Case 1
            ConfigPath = APPPATH & "crashworthiness\Frontal\ConfigFiles\ModerateConfig.xml"
          Case 6
            ConfigPath = APPPATH & "crashworthiness\Frontal\ConfigFiles\SmallOverlapConfig.xml"
          Case 11
            ConfigPath = APPPATH & "crashworthiness\Frontal\ConfigFiles\RightSideSmallOverlapConfig.xml"
          Case 999
            ConfigPath = APPPATH & "crashworthiness\Frontal\ConfigFiles\CustomConfig.xml"
        End Select
        Set Init = Me
      End Function
      
      Public Function GetChannelList(channelGroupName)
          Set GetChannelList = GetNodeNameList(NodeName, channelGroupName)
      End Function
      
      Public Function GetCorridorList(channelGroupName)
          Set GetCorridorList = GetNodeNameList(CorridorNodeName, channelGroupName)
      End Function
        
      Private Function GetNodeNameList(parentNodeName, channelGroupName)
        Dim xDoc : Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")
          Dim settingsGroup : settingsGroup = parentNodeName
          Dim channelSlugList : Set channelSlugList = CreateObject("System.Collections.ArrayList")
          
          If Not xDoc.load(ConfigPath) Then
              Set xDoc = Nothing
              Call LogError("Unable to load crashworthiness configurationfile, config.xml.")
              Exit Function
          End If
          
          Dim childNodeSettingsGroup, childNodeSetting, groupTagValue, groupTag
          
          
          For Each childNodeSettingsGroup In xDoc.documentElement.childNodes
              If childNodeSettingsGroup.nodeName = settingsGroup Then 
              For Each childNodeSetting In childNodeSettingsGroup.childNodes
                  If childNodeSetting.nodeName <> "#comment" Then 
                  groupTag = childNodeSetting.getAttribute("groupTag")
                  If groupTag = channelGroupName Then 
                      channelSlugList.Add(childNodeSetting.nodeName)
                  End If
                  End If
              Next
              End If
          Next  
          Set GetNodeNameList = channelSlugList
          Set xDoc = Nothing
      End Function
      
      Public Function GetConfigValue(settingsGroup, nodeName, attributeName)
          Dim xDoc : Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")
          
          If Not xDoc.load(ConfigPath) Then
              Set xDoc = Nothing
              Call LogError("Unable to load crashworthiness configurationfile, config.xml.")
              Exit Function
          End If
          
          Dim childNodeSettingsGroup, childNodeSetting
          For Each childNodeSettingsGroup In xDoc.documentElement.childNodes
              If childNodeSettingsGroup.nodeName = settingsGroup Then 
              For Each childNodeSetting In childNodeSettingsGroup.childNodes
                  If childNodeSetting.nodeName = nodeName Then 
                  GetConfigValue = childNodeSetting.getAttribute(AttributeName)
                  Exit Function
                  End If
              Next
              End If
          Next  
          
          Set xDoc = Nothing
      End Function
      
      Public Function GetExportList(channelSlug)
        Dim exportList : Set exportList = CreateObject("System.Collections.ArrayList")
        Dim xDoc : Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")
        
        If Not xDoc.load(ConfigPath) Then
            Set xDoc = Nothing
            Call LogError("Unable to load crashworthiness configurationfile, config.xml.")
            Exit Function
        End If
        
        Dim childNodeSettingsGroup, childNodeSetting, exportNode, exportData
        For Each childNodeSettingsGroup In xDoc.documentElement.childNodes
            If childNodeSettingsGroup.nodeName = NodeName Then 
            For Each childNodeSetting In childNodeSettingsGroup.childNodes
                If childNodeSetting.nodeName = channelSlug Then 
                  For Each exportNode In childNodeSetting.childNodes
                    Set exportData = New ExportDTO
                    exportData.Cell = exportNode.getAttribute("cell")
                    exportData.PropertyName = exportNode.getAttribute("property")
                    exportData.RoundValue = exportNode.getAttribute("round")
                    exportData.Divisor = exportNode.getAttribute("divisor")
                    exportData.UseFilteredChannel = cbool(exportNode.getAttribute("useFilteredChannel"))
                    Call exportList.Add(exportData)
                    Set exportData = Nothing
                  Next
                End If
            Next
            End If
        Next  
        
        Set GetExportList = exportList
        Set xDoc = Nothing
      End Function
      Public Function GetDialogExportList
        Dim exportList : Set exportList = CreateObject("System.Collections.ArrayList")
        Dim xDoc : Set xDoc = CreateObject("MSXML2.DOMDocument.6.0")
        
        If Not xDoc.load(ConfigPath) Then
            Set xDoc = Nothing
            Call LogError("Unable to load crashworthiness configurationfile, config.xml.")
            Exit Function
        End If
        
        Dim childNodeSettingsGroup, childNodeSetting, exportNode, exportData
        For Each childNodeSettingsGroup In xDoc.documentElement.childNodes
            If childNodeSettingsGroup.nodeName = DialogNodeName Then 
            For Each childNodeSetting In childNodeSettingsGroup.childNodes
              Set exportData = New ExportDTO
              exportData.Cell = childNodeSetting.getAttribute("cell")
              exportData.PropertyName = childNodeSetting.getAttribute("property")
              Call exportList.Add(exportData)
              Set exportData = Nothing
            Next
            End If
        Next  
        
        Set GetDialogExportList = exportList
        Set xDoc = Nothing
      End Function
    End Class 
    
    Public Function GetConfigFileFactory(testTypeId)
        Set GetConfigFileFactory = (New ConfigFileFactory)(testTypeId)
    End Function
    
    Public Function GetFrontalConfigFileFactory
        Set GetFrontalConfigFileFactory = (New ConfigFileFactory)(1)
    End Function
    
    Class ExportDTO
      Dim Cell, PropertyName, RoundValue, Divisor, UseFilteredChannel
    End Class