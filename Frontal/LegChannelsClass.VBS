'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/23/2017 10:33:07
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ChannelClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\Femur Impulse Calculation.VBS")

Class LegChannels
  Private iGroupName, oGroup, iChannelDictionary, ioccupantLocation, oConfigFileFactory, iGroupTagName, iIsoWhichLeg
  Public Property Get ChannelDictionary
    Set ChannelDictionary = iChannelDictionary
  End Property

  Public Default Function Init(occupantLocation, groupTagName, configFileFactory)
    iGroupTagName = groupTagName
    iOccupantLocation = occupantLocation
    Set oConfigFileFactory = configFileFactory
    Set iChannelDictionary = CreateObject("Scripting.Dictionary")
    iGroupName = oConfigFileFactory.GetConfigValue("testSettings", iGroupTagName, "name")
    iIsoWhichLeg = oConfigFileFactory.GetConfigValue("testSettings", iGroupTagName, "isoWhichLeg")
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
    
    'Femur peak compression time
    Dim femurName : femurName = iChannelDictionary.Item("femurForce").FilteredChannelName
    Dim femurChannel : Set femurChannel = iChannelDictionary.Item("femurForce").FilteredChannel
    Call ChnTMinTMaxCalc("","/" & femurName) 
    Call oGroup.Channels(femurName).Properties.Add("Min_time", mintime*1000, DataTypeFloat64)
    
    Dim compression : Set compression = femurChannel.Properties.Add("MaximumFemurForce", 0, DataTypeFloat64)
    If femurChannel.Minimum < 0 Then
      compression.Value = femurChannel.Minimum
    End If
    Set compression = Nothing
    
    Dim impulseData : Set impulseData = CalculateImpulse(oGroup, femurName)
    Dim kthRiskData : Set kthRiskData = CalculateKTHRisk(oGroup, femurName, impulseData.Impulse)
    Call femurChannel.Properties.Add("Impulse", impulseData.Impulse, DataTypeFloat64)
    Call femurChannel.Properties.Add("KTH_injury", kthRiskData.KTHLocation,DataTypeString)
    Call femurChannel.Properties.Add("KTH_risk", kthRiskData.KTH, DataTypeFloat64)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of femur force failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    Set impulseData = Nothing
    Set kthRiskData = Nothing
    Set femurChannel = Nothing
    
    'Knee displacement time
    Dim kneeDisplacementName : kneeDisplacementName = iChannelDictionary.Item("kneeDeflection").FilteredChannelName
    Dim kneeDisplacementChannel : Set kneeDisplacementChannel = iChannelDictionary.Item("kneeDeflection").FilteredChannel
    Call ChnTMinTMaxCalc("","/" & kneeDisplacementName) 
    Call kneeDisplacementChannel.Properties.Add("Min_time", mintime*1000, DataTypeFloat64)
    
    Dim displacement : Set displacement = kneeDisplacementChannel.Properties.Add("MaximumKneeDisplacement", 0, DataTypeFloat64)
    If kneeDisplacementChannel.Minimum < 0 Then
      displacement.Value = kneeDisplacementChannel.Minimum
    End If
    Set displacement = Nothing
    Set kneeDisplacementChannel = Nothing
    
    If Not Err.number = 0 Then 
      LogError("Calculation of knee displacement failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
      
    'Upper and lower tibia Anterior-Posterior Moment Adjusted (My adjusted)
    Dim apUpName : apUpName = iChannelDictionary.Item("upperTibMomentYAdj").PreFilteredChannelName
    Dim apLoName : apLoName = iChannelDictionary.Item("lowerTibMomentYAdj").PreFilteredChannelName
    Dim tibUpMomentName : tibUpMomentName = iChannelDictionary.Item("upperTibMomentY").FilteredChannelName
    Dim tibLoMomentName : tibLoMomentName = iChannelDictionary.Item("lowerTibMomentY").FilteredChannelName
    Dim tibForceName : tibForceName = iChannelDictionary.Item("lowerTibForceZ").FilteredChannelName
    Call Calculate("Ch(""/" & apUpName & """)=Ch(""/" & tibUpMomentName & """)-Ch(""/" & tibForceName & """)*0.02832",NULL,NULL,"N-m")
    Call Calculate("Ch(""/" & apLoName & """)=Ch(""/" & tibLoMomentName & """)+Ch(""/" & tibForceName & """)*0.006398",NULL,NULL,"N-m")
    Call oGroup.Channels(apUpName).Properties.Add("Test ID", T1, DataTypeString)
    Call oGroup.Channels(apUpName).Properties.Add("Test Type", T4, DataTypeString)
    Call oGroup.Channels(apLoName).Properties.Add("Test ID", T1, DataTypeString)
    Call oGroup.Channels(apLoName).Properties.Add("Test Type", T4, DataTypeString)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of adjusted AP moments failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    'Upper tibia L-M (X moment)time
    'If magnitude to minimum value > max value
    Dim momentXUpName : momentXUpName = iChannelDictionary.Item("upperTibMomentX").FilteredChannelName
    Call SetAbsMaxAndTime(oGroup, momentXUpName, "Max_moment", "Max_time")
    
    'Lower tibia L-M (X moment)time
    'If magnitude to minimum value > max value
    Dim momentXLoName : momentXLoName = iChannelDictionary.Item("lowerTibMomentX").FilteredChannelName
    Call SetAbsMaxAndTime(oGroup, momentXLoName, "Max_moment", "Max_time")
    
    'Upper tibia A-P (Y moment)time
    'If magnitude to minimum value > max value
    Call SetAbsMaxAndTime(oGroup, apUpName, "Max_moment", "Max_time")
      
    'Lower tibia A-P (Y moment)time
    'If magnitude to minimum value > max value
    Call SetAbsMaxAndTime(oGroup, apLoName, "Max_moment", "Max_time")
      
    'Calculate tibia resultant moments (left and upper/lower)
    Dim momentRUpName : momentRUpName = iChannelDictionary.Item("upperTibMomentR").PreFilteredChannelName
    Dim momentRLoName : momentRLoName = iChannelDictionary.Item("lowerTibMomentR").PreFilteredChannelName
    
    Call Calculate("Ch(""/" & momentRUpName & """)=sqrt(Ch(""/" & momentXUpName & """)^2+Ch(""/" & apUpName & """)^2)",NULL,NULL,"N-m") 
    Call ChnTMinTMaxCalc("","/" & momentRUpName)
    Call oGroup.Channels(momentRUpName).Properties.Add("Max_time", maxtime*1000, DataTypeFloat64)
    If Not Err.number = 0 Then 
      LogError("Calculation of upper tibia resultant moment failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If     
    
    Call Calculate("Ch(""/" & momentRLoName & """)=sqrt(Ch(""/" & momentXLoName & """)^2+Ch(""/" & apLoName & """)^2)",NULL,NULL,"N-m") 
    Call ChnTMinTMaxCalc("","/" & momentRLoName)
    Call oGroup.Channels(momentRLoName).Properties.Add("Max_time", maxtime*1000, DataTypeFloat64)
    If Not Err.number = 0 Then 
      LogError("Calculation of lower tibia resultant moment failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If 
    
    'Lower tibia axial force time
    Call ChnTMinTMaxCalc("","/" & tibForceName)
    Call oGroup.Channels(tibForceName).Properties.Add("Min_time", mintime*1000, DataTypeFloat64)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of lower tibia force failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    'Tibia indices (left upper/lower)
    Dim posResultName : posResultName = "PosResult"
    Dim negResultName : negResultName = "NegResult"
    Call ChnSplitAtValue("/" & tibForceName,0,0,"/" & posResultName,"/" & negResultName) 
    
    Dim tIUpName : tIUpName = iChannelDictionary.Item("upperTI").PreFilteredChannelName
    Dim tILoName : tILoName = iChannelDictionary.Item("lowerTI").PreFilteredChannelName
    Call Calculate("Ch(""/" & tIUpName & """)=(Ch(""/" & momentRUpName & """)/225)+Abs(Ch(""/" & negResultName & """)/35900)",NULL,NULL,"") 
    Call Calculate("Ch(""/" & tILoName & """)=(Ch(""/" & momentRLoName & """)/225)+Abs(Ch(""/" & negResultName & """)/35900)",NULL,NULL,"") 
    
    Call oGroup.Channels.Remove(negResultName)
    Call oGroup.Channels.Remove(posResultName)
    
    Call ChnTMinTMaxCalc("","/" & tIUpName)
    Call oGroup.Channels(tIUpName).Properties.Add("Max_time", maxtime*1000, DataTypeFloat64)
    Call ChnTMinTMaxCalc("","/" & tILoName)  
    Call oGroup.Channels(tILoName).Properties.Add("Max_time", maxtime*1000, DataTypeFloat64)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of TIs failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    'Foot X acceleration time
    Dim footX : footX = iChannelDictionary.Item("footX").FilteredChannelName
    Call SetAbsMaxAndTime(oGroup, footX, "Max_acceleration", "Max_time")
    
    'Foot Z acceleration time
    Dim footZ : footZ = iChannelDictionary.Item("footZ").FilteredChannelName
    Call SetAbsMaxAndTime(oGroup, footZ, "Max_acceleration", "Max_time")

    'Foot resultant acceleration & max time
    Dim footR : footR = iChannelDictionary.Item("footR").PreFilteredChannelName
    Call Calculate("Ch(""/" & footR & """) = SQRT((Ch(""/" & footX & """))^2+(Ch(""/" & footZ & """))^2)",NULL,NULL,"gn") 
    Call ChnTMinTMaxCalc("","/" & footR) 
    Call oGroup.Channels(footR).Properties.Add("Test ID", T1, DataTypeString)
    Call oGroup.Channels(footR).Properties.Add("Test Type", T4, DataTypeString)
    Call oGroup.Channels(footR).Properties.Add("Max_time", maxtime*1000, DataTypeFloat64)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of " & iIsoWhichLeg & " foot resultant failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    On Error Goto 0 
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