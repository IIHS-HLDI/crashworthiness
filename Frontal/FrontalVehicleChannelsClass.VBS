'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/23/2017 11:34:17
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ChannelClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")

Class FrontalVehicleChannels
  Private iGroupName, oGroup, iChannelDictionary, iVehicleNumber, oConfigFileFactory, iGroupTagName
  Public Property Get ChannelDictionary
    Set ChannelDictionary = iChannelDictionary
  End Property

  Public Default Function Init(vehicleNumber, testTypeId)
    iGroupTagName = "vehicleGroupName"
    iVehicleNumber = vehicleNumber
    Set oConfigFileFactory = GetConfigFileFactory(testTypeId)
    Set iChannelDictionary = CreateObject("Scripting.Dictionary")
    iGroupName = oConfigFileFactory.GetConfigValue("testSettings", iGroupTagName, "name")
    Set oGroup = CreateGroup(iVehicleNumber & iGroupName)
    Call LoadChannels
    Set Init = Me
  End Function
 
  Private Sub LoadChannels
    Dim slugs : Set slugs = oConfigFileFactory.GetChannelList(iGroupTagName)
    Dim slug, key, position
    For Each slug in slugs
      key = oConfigFileFactory.GetConfigValue("ChannelMetaData", slug, "key")
      position = oConfigFileFactory.GetConfigValue("ChannelMetaData", slug, "position")
      If Not IsEmpty(position) Then
        Call iChannelDictionary.Add(key, (New Channel)(iVehicleNumber & position, slug, oConfigFileFactory))
      Else
        Call LogError("Position attribute not found in config file for: " & slug & ". Position information is required for vehicle mounted sensors")
      End If
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

   'Vehicle acceleration and delta V
    Dim vehX : Set vehX = iChannelDictionary.Item("vehicleX").FilteredChannel
    Dim vehY : Set vehY = iChannelDictionary.Item("vehicleY").FilteredChannel
    Dim vehZ : Set vehZ = iChannelDictionary.Item("vehicleZ").FilteredChannel
    
    Dim vehR
    Dim vehRName : vehRName = iChannelDictionary.Item("vehicleR").PreFilteredChannelName
    If oGroup.Channels.Exists(vehRName) Then
      Set vehR = oGroup.Channels(vehRName)
    Else
      Set vehR = oGroup.Channels.Add(vehRName, DataTypeChnFloat64)  
    End If
        
    Call ChnXYZAbsValue(vehX, vehY, vehZ, vehR)
    
    Call ChnTMinTMaxCalc("","/" & vehRName) 
    Call vehR.Properties.Add("Test ID", T1, DataTypeString)
    Call vehR.Properties.Add("Test Type", T4, DataTypeString)
    Call vehR.Properties.Add("Max_time", maxtime*1000, DataTypeFloat64)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of vehicle resultant acceleration failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
      
    'Vehicle X acceleration time
    Call SetAbsMaxAndTime(oGroup, vehX.Name, "Max_acceleration", "Max_time")
    
    'Vehicle Y acceleration time
    Call SetAbsMaxAndTime(oGroup, vehY.Name, "Max_acceleration", "Max_time")

    'Vehicle Z acceleration time
    Call SetAbsMaxAndTime(oGroup, vehZ.Name, "Max_acceleration", "Max_time")
    
    Set vehX = Nothing
    Set vehY = Nothing
    Set vehZ = Nothing
    Set vehR = Nothing
    
    Dim vehicleX : Set vehicleX = iChannelDictionary.Item("vehicleX")
    Dim vehicleXFilterCName : vehicleXFilterCName = "VehicleXCFC180"
    Call ChnCFCFiltCalc("", vehicleX.PreFilteredChannel, "/" & vehicleXFilterCName,"CFC_180",0,"EndPoints",10)
    Dim vehicleDeltaVName : vehicleDeltaVName = "VEHC0000__VEX"
    Dim impactVelocity : impactVelocity = data.Root.Properties("Impact_velocity").Value
    'Calculate delta-V
    'Remove data before time zero and create new channel then integrate data
    Call ChnTimeAreaCopy("","/" & vehicleXFilterCName,"","/TimeRangeCopyY",0,1)
    Call ChnIntegrate("","/TimeRangeCopyY","/Integrated") 
    'Take integrated acceleration and multiply by 35.316 (9.81*3.6) to convert from m/s to km/h 
    'T5 is the impact velocity
    Call Calculate("Ch(""/" & vehicleDeltaVName & """)=Ch(""/Integrated"")*9.81*3.6+" & impactVelocity,NULL,NULL,"km/h")
    oGroup.Channels(vehicleDeltaVName).Properties("wf_xname").Value = "Time_VEHC0000__VEX"
    
    'Delete channels that aren't needed
    Call oGroup.Channels.Remove("TimeRangeCopyY")
    Call oGroup.Channels.Remove("Integrated")
    Call oGroup.Channels.Remove(vehicleXFilterCName)
    
    If Not Err.number = 0 Then 
      LogError("Calculation of vehicle delta v failed. Message: " & Replace(Err.Description, vbCrLf, ""))
      Err.Clear
    End If
    
    On Error Goto 0
    Set vehicleX = Nothing
  End Sub
  
  Public Sub ExportProperties(workbook)
    Dim sheetName : sheetName = oConfigFileFactory.GetConfigValue("testSettings", "exportSheets", "occupant")
    sheetName = "11" & sheetName
    Dim sheet : Set sheet = workbook.Sheets(sheetName)
    sheet.Select
    Dim item, exportData, chn
    For Each item In iChannelDictionary.Items
      item.ExportProperties(sheet)
    Next
    
    Set chn = Nothing
  End Sub
End Class