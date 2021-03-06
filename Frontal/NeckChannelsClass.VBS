'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/22/2017 17:28:54
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ChannelClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")

Class NeckChannels
  Private iGroupName, oGroup, iChannelDictionary, ioccupantLocation, oConfigFileFactory, iGroupTagName
  Public Property Get ChannelDictionary
    Set ChannelDictionary = iChannelDictionary
  End Property

  Public Default Function Init(occupantLocation, configFileFactory)
    iGroupTagName = "neckGroupName"
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
   
    'Neck injury calculations
    Dim oNeckShear : Set oNeckShear = iChannelDictionary.Item("neckForceX")
    
    'Neck shear force
    Call SetAbsMaxAndTime(oGroup, oNeckShear.FilteredChannelName, "Max_shear_force", "Max_time")
    
    'Neck lateral force
    Dim oNeckLateral : Set oNeckLateral = iChannelDictionary.Item("neckForceY")
    
    Call SetAbsMaxAndTime(oGroup, oNeckLateral.FilteredChannelName, "LateralNeckForce", "LateralNeckForce_time")
    
    'Neck z force 
    Dim oNeckZ : Set oNeckZ = iChannelDictionary.Item("neckForceZ")
    Call ChnTMinTMaxCalc("", oNeckZ.FilteredChannel)
    Call oNeckZ.FilteredChannel.Properties.Add("Min_time", mintime*1000, DataTypeFloat64)
    Call oNeckZ.FilteredChannel.Properties.Add("Max_time", maxtime*1000, DataTypeFloat64)
    
    'Neck tension
    Dim neckTension : Set neckTension = oNeckZ.FilteredChannel.Properties.Add("NeckTension", 0, DataTypeFloat64)
    If CMax(oNeckZ.FilteredChannel) > 0 Then
      neckTension.Value = CMax(oNeckZ.FilteredChannel)
    End If
    Set neckTension = Nothing
    
    'Neck compression
    Dim neckCompression : Set neckCompression = oNeckZ.FilteredChannel.Properties.Add("NeckCompression", 0, DataTypeFloat64)
    If CMin(oNeckZ.FilteredChannel) < 0 Then
      neckCompression.Value = Abs(CMin(oNeckZ.FilteredChannel))
    End If
    Set neckCompression = Nothing
    
    If Not Err.number = 0 Then 
      LogError("Calculation of neck tension/compression failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    'Calculate MOCy & Flextion (max) and Extension (min) time
    Dim MOCyName : MOCyName = iChannelDictionary.Item("neckMomentYOc").PreFilteredChannelName
    Dim MomentYName : MomentYName = iChannelDictionary.Item("neckMomentY").FilteredChannelName
    Dim neckShearFilterBName : neckShearFilterBName = "NeckShearCFC600"
    Call ChnCFCFiltCalc("", oNeckShear.PreFilteredChannel, "/" & neckShearFilterBName,"CFC_600",0,"EndPoints",10)
    Call Calculate("Ch(""/" & MOCyName & """)=Ch(""/" & MomentYName & """)-Ch(""/" & neckShearFilterBName & """)*0.01778",NULL,NULL,"N-m") 
    Call ChnTMinTMaxCalc("","/" & MOCyName) 
    Call oGroup.Channels.Remove(neckShearFilterBName)
    Call oGroup.Channels(MOCyName).Properties.Add("Max_time", maxtime*1000, DataTypeFloat64)
    Call oGroup.Channels(MOCyName).Properties.Add("Min_time", mintime*1000, DataTypeFloat64)
    
    Dim momentOc : Set momentOc = iChannelDictionary.Item("neckMomentYOc").PreFilteredChannel
    Dim flexion : Set flexion = momentOc.Properties.Add("NeckFlexion", 0, DataTypeFloat64)
    
    If momentOc.Maximum > 0 Then
      flexion.Value = momentOc.Maximum
    End If
    Set flexion = Nothing
    
    Dim extension : Set extension = momentOc.Properties.Add("NeckExtension", 0, DataTypeFloat64)
    If momentOc.Minimum < 0 Then
      extension.Value = Abs(momentOc.Minimum)
    End If
    Set extension = Nothing
    Set momentOc = Nothing
    
    If Not Err.number = 0 Then 
      LogError("Calculation of neck oc moment failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If        
            
    'Nij calculattions
    Dim neckZFilterBName : neckZFilterBName = "NeckZCFC600"
    Call ChnCFCFiltCalc("", oNeckZ.PreFilteredChannel, "/" & neckZFilterBName,"CFC_600",0,"EndPoints",10)
    
    Dim nteName : nteName = iChannelDictionary.Item("nte").PreFilteredChannelName
    Call ChnNijCalc("","/" & neckZFilterBName,"/" & MOCyName,"/" & nteName,6806,-135,"NTE")
    Dim nte : Set nte = ogroup.Channels(nteName)
    Call nte.Properties.Add("NTE Time", NIJTime*1000, DataTypeFloat64)
    
    Dim ntfName : ntfName = iChannelDictionary.Item("ntf").PreFilteredChannelName
    Call ChnNijCalc("","/" & neckZFilterBName,"/" & MOCyName,"/" & ntfName,6806,310,"NTF")
    Dim ntf : Set ntf = ogroup.Channels(ntfName)
    Call ntf.Properties.Add("NTF Time", NIJTime*1000, DataTypeFloat64)
    
    Dim nceName : nceName = iChannelDictionary.Item("nce").PreFilteredChannelName
    Call ChnNijCalc("","/" & neckZFilterBName,"/" & MOCyName,"/" & nceName,-6160,-135,"NCE")
    Dim nce : Set nce = ogroup.Channels(nceName)
    Call nce.Properties.Add("NCE Time", NIJTime*1000, DataTypeFloat64)
    
    Dim ncfName : ncfName = iChannelDictionary.Item("ncf").PreFilteredChannelName
    Call ChnNijCalc("","/" & neckZFilterBName,"/" & MOCyName,"/" & ncfName,-6160,310,"NCF")
    Dim ncf : Set ncf = ogroup.Channels(ncfName)
    Call ncf.Properties.Add("NCF Time", NIJTime*1000, DataTypeFloat64)
    Call oGroup.Channels.Remove(neckZFilterBName)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of nij failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    On Error Goto 0
    
    Set oNeckShear = Nothing
    Set oNeckLateral = Nothing
    Set oNeckZ = Nothing
    Set ntf = Nothing
    Set ncf = Nothing
    Set ncf = Nothing
    Set nce = Nothing 
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