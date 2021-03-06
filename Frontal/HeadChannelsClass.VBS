'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 08/02/2016 17:55:47
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ChannelClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")

Class HeadChannels
  Private iHeadX, iHeadY, iHeadZ, iGroupName, oGroup, iChannelDictionary, ioccupantLocation, oConfigFileFactory, iGroupTagName
  Public Property Get ChannelDictionary
    Set ChannelDictionary = iChannelDictionary
  End Property

  Public Default Function Init(occupantLocation, configFileFactory)
    iGroupTagName = "headGroupName"
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
    
    Dim headX : Set headX = iChannelDictionary.Item("headX").FilteredChannel
    Dim headY : Set headY = iChannelDictionary.Item("headY").FilteredChannel
    Dim headZ : Set headZ = iChannelDictionary.Item("headZ").FilteredChannel
    
    'Peak Head Resultant Acceleration Time & HIC Values
    Dim headR
    Dim headRName : headRName = iChannelDictionary.Item("headR").PreFilteredChannelName
    If oGroup.Channels.Exists(headRName) Then
      Set headR = oGroup.Channels(headRName)
    Else
      Set headR = oGroup.Channels.Add(headRName, DataTypeChnFloat64)  
    End If
    
    Call ChnXYZAbsValue(headX, headY, headZ, headR)
    If Not Err.number = 0 Then 
      LogError("Calculation of resultant head acceleration failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    'Calculate HIC
    Call ChnHICCalc("",headR,1,1,0,0,10,0)
    Call ChnTMinTMaxCalc("",headR)   
    Call headR.Properties.Add("Max_time", maxtime*1000, DataTypeFloat64) 
    Call headR.Properties.Add("HIC36", HICRes(1,2), DataTypeFloat64)
    Call headR.Properties.Add("HIC36_T1",HICRes(2,2), DataTypeFloat64)
    Call headR.Properties.Add("HIC36_T2", HICRes(3,2), DataTypeFloat64)
    Call headR.Properties.Add("HIC15", HICRes(1,3), DataTypeFloat64)
    Call headR.Properties.Add("HIC15_T1", HICRes(2,3), DataTypeFloat64)
    Call headR.Properties.Add("HIC15_T2", HICRes(3,3), DataTypeFloat64)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of HIC failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
      
    'Calculate 3ms clip
    Call ChnXMSCalc("",headR,3,0,"in one peak",0)
    Call headR.Properties.Add("Clip_3_ms", XmsAcceleration, DataTypeFloat64)
    Call headR.Properties.Add("Clip_3_ms_T1", XmsTimeBegin, DataTypeFloat64)
    Call headR.Properties.Add("Clip_3_ms_T2",XmsTimeEnd , DataTypeFloat64)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of 3ms clip failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
      
    'Integrate and remove offset for angular rate sensors
    Dim headArsX : Set headArsX = iChannelDictionary.Item("headArsX").FilteredChannel
    Dim headArsY : Set headArsY = iChannelDictionary.Item("headArsY").FilteredChannel
    Dim headArsZ : Set headArsZ = iChannelDictionary.Item("headArsZ").FilteredChannel
    
    Dim oHeadAnX : Set oHeadAnX = iChannelDictionary.Item("headAnX")
    Dim oHeadAnY : Set oHeadAnY = iChannelDictionary.Item("headAnY")
    Dim oHeadAnZ : Set oHeadAnZ = iChannelDictionary.Item("headAnZ")
        
    Dim arsArray : arsArray = Array(headArsX, headArsY, headArsZ)
    Dim angleArray : angleArray = Array(oHeadAnX, oHeadAnY, oHeadAnZ)
    Dim i
    For i = 0 To Ubound(arsArray)
      Call ChnIntegrate("", arsArray(i), "/" & angleArray(i).PreFilteredChannelName) 
      Call ChnOffset(angleArray(i).PreFilteredChannel, angleArray(i).PreFilteredChannel, -angleArray(i).PreFilteredChannel.Values(1001),"free offset")
      Call arsArray(i).Properties.Add("AbsMax", MaxValue(arsArray(i)))
      Call arsArray(i).Properties.Add("AbsMax_time", AbsMaxTime(arsArray(i)))
    Next
    
    If Not Err.number = 0 Then 
      LogError("Calculation of ARS channels failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    'Add BRIC to resultand head acceleration channel
    Call headR.Properties.Add("BRIC", BRICcalc(headArsX.Properties("AbsMax").Value, headArsY.Properties("AbsMax").Value, headArsZ.Properties("AbsMax").Value))
    
    If Not Err.number = 0 Then 
      LogError("Calculation of BRIC failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    On Error Goto 0 
    Set headX = Nothing
    Set headY = Nothing
    Set headZ = Nothing
    Set headR = Nothing
    Set oHeadAnX = Nothing
    Set oHeadAnY = Nothing
    Set oHeadAnZ = Nothing
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