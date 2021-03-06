'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/23/2017 14:28:08
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ChannelClass.VBS")

Class Corridor
  Private iIsCalculated, iIsoCode, iSourceChannelTag, iSourceChannelGroupTag, iOccupantLocation, iDescription, iGroupTag, iChannelSlug,  iCorridorChannelName, iReferenceTag, iInvertSource, oConfigFileFactory, oGroup
  
  Public Property Get CorridorChannel
    If oGroup.Channels.Exists(iCorridorChannelName) Then
        Set CorridorChannel = oGroup.Channels(iCorridorChannelName)
    Else
        Set CorridorChannel = Nothing
    End If
  End Property 
  
  Public Property Get InvertSource
    If Not IsEmpty(iInvertSource) Then InvertSource = cbool(iInvertSource)
  End Property
  
  Public Property Get IsCalculated
    If Not IsEmpty(IisCalculated) Then IsCalculated = cbool(IisCalculated)
  End Property
  
  Public Property Get ReferenceChannel
    Dim chn : Set chn = (New Corridor)(iOccupantLocation, iReferenceTag, oConfigFileFactory)
    Set ReferenceChannel = chn.CorridorChannel
    Set chn = Nothing
  End Property 
  
  Public Property Get Description
    Description = iDescription
  End Property
  
  Public Property Get SourceChannel
    Dim grpName
    If Not IsEmpty(iSourceChannelTag) And Not IsEmpty(iSourceChannelGroupTag) Then 
      grpName = iOccupantLocation & oConfigFileFactory.GetConfigValue("testSettings", iSourceChannelGroupTag, "name")
      If Data.Root.ChannelGroups.Exists(grpName) Then
        Dim chn : Set chn = (New Channel)(iOccupantLocation, iSourceChannelTag, oConfigFileFactory)
        Set SourceChannel = chn.FilteredChannel
      End If
    End If   
    Set chn = Nothing
  End Property 

  Public Property Get CorridorChannelName
    CorridorChannelName = iCorridorChannelName
  End Property
  
  Public Default Function Init(occupantLocation, channelSlug, configFileFactory)
    iOccupantLocation = occupantLocation
    iChannelSlug = channelSlug
    Set oConfigFileFactory = configFileFactory
    Call GetChannelInfo
    Set oGroup = Data.Root.ChannelGroups(iOccupantLocation & oConfigFileFactory.GetConfigValue("testSettings", iGroupTag, "name"))
    Set Init = Me
  End Function
  
  Private Sub GetChannelInfo
    Dim corridorsTagName : corridorsTagName = "CorridorsMetaData"
    iIsCalculated = oConfigFileFactory.GetConfigValue(corridorsTagName, iChannelSlug, "isCalculated")
    iIsoCode = oConfigFileFactory.GetConfigValue(corridorsTagName, iChannelSlug, "isoCode")
    iSourceChannelTag = oConfigFileFactory.GetConfigValue(corridorsTagName, iChannelSlug, "sourceChannelTag")
    iSourceChannelGroupTag = oConfigFileFactory.GetConfigValue(corridorsTagName, iChannelSlug, "sourceChannelGroupTag")
    iReferenceTag = oConfigFileFactory.GetConfigValue(corridorsTagName, iChannelSlug, "referenceTag")
    iInvertSource = oConfigFileFactory.GetConfigValue(corridorsTagName, iChannelSlug, "invertSource")
    iDescription = oConfigFileFactory.GetConfigValue(corridorsTagName, iChannelSlug, "description")
    iGroupTag = oConfigFileFactory.GetConfigValue(corridorsTagName, iChannelSlug, "groupTag")
    iCorridorChannelName = iOccupantLocation & iIsoCode
  End Sub
End Class