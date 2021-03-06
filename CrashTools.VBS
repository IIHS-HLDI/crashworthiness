'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 08/02/2016 17:55:47
'-- Author: 
'-- Comment: Must be placed in root directory
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

Function GetTestIdPrefix(TestId)
  Dim Length, i, Prefix, Char 
  Length = Len(TestId)
  For i = 1 to Length
    Char = Mid(TestId, i, 1)
    If IsNumeric(Char) = False Then
      Prefix = Prefix & Char
    Else
      GetTestIdPrefix = Prefix
      Exit Function
    End If
  Next
  Call MsgBox("Could not determine IIHSTestId prefix")
  GetTestIdPrefix = Null
End Function

Function GetTestTypeId(iihsTestId)
  Dim id : id = LCase(iihsTestId)
  
  If InStr(id, "f") > 0 Then 
    GetTestTypeId = 1 'moderate
  ElseIf InStr(id, "n") > 0 Then 
    GetTestTypeId = 6 'small overlap
  ElseIf InStr(id, "p") > 0 Then 
    GetTestTypeId = 11 'right side small overlap
  ElseIf InStr(id, "s") > 0 Then 
    GetTestTypeId = 2 'side impact
  Else
    GetTestTypeId = Empty
  End If 
  
End Function

Function IsConnected
  Dim Filesys, Newfolder
  Set Filesys = CreateObject("Scripting.FileSystemObject")
  If Not Filesys.FolderExists(CurrentScriptPath) THEN
    IsConnected = False
    Exit Function
  Else
    IsConnected = True
    Exit Function
  End If
End Function

Function IsDataFileLoaded
  If Data.Root.Name = "NoName" Then
    IsDataFileLoaded = False
    Exit Function
  Else
    IsDataFileLoaded = True
    Exit Function
  End If
End Function

Function IsViewLoaded
  If View.FileName = "NONAME" Then
    IsViewLoaded = False
    Exit Function
  Else
    IsViewLoaded = True
    Exit Function
  End If
End Function

Function GetTestDirectory (TestID)
  dim fso, RootPath, folder, TestType, TestYear
  
  set fso = CreateObject("Scripting.FileSystemObject")
  If Left(TestID,3) = "CEF" or Left(TestID,3) = "CES" or Left(TestID,3) = "CEN" or Left(TestID,3) = "CEP" or Left(TestID,3) = "SER" Then
    TestType = "CE"
  ElseIf Left(TestID,2) = "CF" or Left(TestID,2) = "CS" or Left(TestID,2) = "CN" or Left(TestID,2) = "CP"  or Left(TestID,2) = "SR"Then
    TestType = "R&D"
  Else
    Exit Function
  End If
  
  If Left(TestID, 2) = "CS" or Left(TestID, 3) = "CES" Then
    RootPath = "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\Side-Impact Crash Tests\"
  ElseIf Left(TestID, 2) = "CF" or Left(TestID, 3) = "CEF" or  Left(TestID, 2) = "CN" or Left(TestID, 3) = "CEN" or  Left(TestID, 2) = "CP" or Left(TestID, 3) = "CEP" Then
    RootPath = "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\Frontal Crash Tests\"
  ElseIf Left(TestID, 2) = "SR" or Left(TestID, 3) = "SER"  Then
    RootPath = "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\Rear Impact Sled Tests\"
  End If
    
  For Each folder in fso.GetFolder(RootPath).SubFolders
      If folder.Name = TestID Then
        GetTestDirectory = folder.Path
      End If
  Next
  
  For Each folder in fso.GetFolder(RootPath & "Archived\").SubFolders
    If TestType = "CE" Then 
      If Mid(TestID, 4, 2) = right(folder.Name, 2) Then 
        RootPath = RootPath & "Archived\" & folder.Name & "\" & TestType & "\"
      End If
    End If
    If TestType = "R&D" Then
      If Mid(TestID, 3, 2) = right(folder.Name, 2) Then
        RootPath = RootPath & "Archived\" & folder.Name & "\" & TestType & "\"
      End If
    End If 
  Next
  
  For Each folder in fso.GetFolder(RootPath).SubFolders
      If folder.Name = TestID Then
        GetTestDirectory = folder.Path
        Exit Function
      End If
  Next
End Function

Sub TransferPropToCellRef (ExcelSheet, Row, Column, PropLocationEnum, ChannelGroup, ChannelName, PropertyName, Precision, Divisor)  
  Dim oProperty : Set oProperty = GetProperty(PropLocationEnum, ChannelGroup, ChannelName, PropertyName)
  If oProperty Is Nothing Then Exit Sub
  
  If Not IsNull(oProperty.Value) Then
    Call SetExcelProp(ExcelSheet.Cells(row, column), oProperty, Precision, Divisor)
  Else
    Call LogError("Property value is null. Info: ChannelGroupName: " & ChannelGroup & ", ChannelName: " & ChannelName & ", PropertyName: " & PropertyName)
  End If 
  Set oProperty = Nothing
End Sub

Sub TransferPropToNamedCell(ExcelSheet, CellName, PropLocationEnum, ChannelGroup, ChannelName, PropertyName, Precision, Divisor)  
  Dim oProperty : Set oProperty = GetProperty(PropLocationEnum, ChannelGroup, ChannelName, PropertyName)
  If oProperty Is Nothing Then Exit Sub
  
  If Not IsNull(oProperty.Value) Then
    Call SetExcelProp(ExcelSheet.Range(CellName), oProperty, Precision, Divisor)
  Else
    Call LogError("Property value is null. Info: ChannelGroupName: " & ChannelGroup & ", ChannelName: " & ChannelName & ", PropertyName: " & PropertyName)
  End If 
  Set oProperty = Nothing
End Sub

Sub SetExcelProp(oExcelProp, oProperty, Precision, Divisor)
  Select Case oProperty.DataType
    Case DataTypeFloat32, DataTypeFloat64
      oExcelProp.Value = round((oProperty.Value)/Divisor, Precision)
    Case Else
      oExcelProp.Value = oProperty.Value
  End Select  
End Sub

Sub TransferPropToExcel(ExcelSheet, CellName, oParent, propertyName, Precision, Divisor)
  Dim oProperty
  
  On Error Resume Next
    Set oProperty = oParent.Properties(propertyName)
  If Err.number <> 0 Then
    Call LogError("Unable to find " & propertyName & "  property to export. Info: " & Err.Description)
    Set oProperty = Nothing
  End If  
  
  On Error Goto 0
  If oProperty Is Nothing Then Exit Sub
  
  If Not IsNull(oProperty.Value) Then
    Select Case oProperty.DataType
      Case DataTypeFloat32, DataTypeFloat64
        ExcelSheet.Range(CellName).Value = round((oProperty.Value)/Divisor,Precision)
      Case Else
        ExcelSheet.Range(CellName).Value = oProperty.Value
    End Select  
  Else
    Call LogError("Property value is null. Info: ChannelGroupName: " & ChannelGroup & ", ChannelName: " & ChannelName & ", PropertyName: " & PropertyName)
  End If 
End Sub

Function GetProperty(PropLocationEnum, ChannelGroup, ChannelName, PropertyName)
  Dim oProperty
  
  On Error Resume Next
  Select Case PropLocationEnum
    Case eSearchFile
      Set oProperty = data.Root.Properties(PropertyName)
    Case eSearchChannelGroup
      Set oProperty = data.Root.ChannelGroups(ChannelGroup).Properties(PropertyName)
    Case eSearchChannel
      Set oProperty = data.Root.ChannelGroups(ChannelGroup).Channels(ChannelName).Properties(PropertyName)
  End Select
  
  If Err.number <> 0 Then
    Call LogError("Unable to find property to export. Info: ChannelGroupName: " & ChannelGroup & ", ChannelName: " & ChannelName & ", PropertyName: " & PropertyName)
    On Error Goto 0
    Set GetProperty = Nothing
    Exit Function
  End If
  
  On Error Goto 0
  Set GetProperty = oProperty
End Function

Function GetTestErrorLog
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  Dim testRoot : testRoot = Data.Root.Properties("data_path").Value
  Dim testId : testId = data.Root.Properties("Test_ID").Value
  Dim errorLogPath : errorLogPath = testRoot & "\DATA\EXCEL\" & testId & " Error Log.txt"
  
  If Not fso.FileExists(errorLogPath) Then
    fso.CreateTextFile(errorLogPath).Close
  End If
  
  Set GetTestErrorLog = fso.GetFile(errorLogPath)
  Set fso = Nothing
End Function

Function LastErrorCount
  Dim file : Set file = GetTestErrorLog
  Dim stream : Set stream = file.OpenAsTextStream(1, -2)
  Dim i
  Do While Not stream.AtEndOfStream
    If Right(stream.ReadLine, 10) = "Started..." Then
      i = 0
    End If
    i = i + 1
  Loop
  stream.Close
  Set stream = Nothing
  Set file = Nothing
  LastErrorCount = i - 2
End Function

Sub LaunchLog
  Dim file : Set file = GetTestErrorLog
  Call Extprogram("notepad.exe", file.Path)
  Set file = Nothing
End Sub

Sub LogError(ErrorMessage)
  Dim fso : Set fso = createobject("scripting.FileSystemObject")
  Dim file : Set file = GetTestErrorLog
  Dim textStream : Set textStream = file.OpenAsTextStream(8, -2)
  
  textStream.WriteLine(Now & ": " & ErrorMessage)
  textStream.Close
  
  Set fso = Nothing
  Set file = Nothing
  Set textStream = Nothing
End Sub

Function MaxValue(Channel)
  If Channel.Properties("maximum").Value >= ABS(Channel.Properties("minimum").value) Then
    MaxValue = round(Channel.Properties("maximum").Value,1)
  Else
    MaxValue = round(Channel.Properties("minimum").Value,1)
  End If
End Function
  
Function AbsMaxTime(Channel)
  If Channel.Properties("maximum").Value >= ABS(Channel.Properties("minimum").value) Then
    Call ChnTMinTMaxCalc("","/" & Channel.Name)
    AbsMaxTime = round(maxtime*1000,1)
  Else
    Call ChnTMinTMaxCalc("","/" & Channel.Name)
    AbsMaxTime = round(mintime*1000,1)
  End If
End Function

Function BRICcalc(Xmax, Ymax, Zmax)
  BRICcalc = SQRT((Xmax/3795.84)^2+(Ymax/3234.34)^2+(Zmax/2456.27)^2)
End Function

Function GetRawGroupName(channelGroupName)
  Dim groups : Set groups = Data.GetChannelGroups("*" & channelGroupName)
  If groups.Count <> 1 Then 
    Call Err.Raise(vbObjectError + 10, "IIHS Crash Test Processing", "There is either 0 or more than 1 raw channel group.")
  Else
    GetRawGroupName = groups.Item(1).Name
  End If
End Function

Function CreateGroup(ChannelGroupName)
  If Data.Root.ChannelGroups.Exists(ChannelGroupName) Then
    Set CreateGroup = Data.Root.ChannelGroups(ChannelGroupName)
  Else
    Set CreateGroup = Data.Root.ChannelGroups.Add(ChannelGroupName)
  End If
End Function

Sub SetAbsMaxAndTime(Group, ChannelName, MaxValPropName, TimeOfMaxPropName)
    Dim maxVal, timeOfMax
    Dim Channel : Set Channel = Group.Channels(ChannelName)
    
    Call ChnTMinTMaxCalc("", Channel)
    If abs(Channel.Minimum) > Channel.Maximum Then
      maxVal = channel.Minimum
      timeOfMax = mintime*1000
    Else
      maxVal = channel.Maximum
      timeOfMax = maxtime*1000
    End If
    Call Channel.Properties.Add(MaxValPropName, maxVal, DataTypeFloat64)
    Call Channel.Properties.Add(TimeOfMaxPropName, timeOfMax, DataTypeFloat64)
    Set Channel = Nothing
    If Not Err.number = 0 Then 
      LogError("Calculation of " & channelname & ":" & MaxValPropName & " failed. Message: " & Err.Description)
      Err.Clear
    End If
  End Sub
  
  Function GetWorkbook(path)
    Const xlMaximized =&HFFFFEFD7
    Dim oExcel : Set oExcel = CreateObject("Excel.Application")
    Dim oWorkbook : Set oWorkbook = oExcel.Workbooks.Open(path) 
    oExcel.WindowState = xlMaximized
    oExcel.Visible = true 
    Set GetWorkbook = oWorkbook  
  End Function
  
  Sub PreFilterRawData(oGroup, iihsTestId, testDescription)
    Dim Chn

    For Each Chn In oGroup.Channels
      If Chn.Name <> "Time axis" Then
        Call ChnToWfChn("[1]/Time axis",Chn,0, "WfXRelative")
        Call ChnTimeAreaCopy("", Chn, "", Chn, -0.1, 0.5)
        Call Chn.Properties.Add("Test ID", iihsTestId, DataTypeString)
        Call Chn.Properties.Add("Test Type", testDescription, DataTypeString)
        If Chn.Properties("unit_string").Value = "G" Then
            Chn.Properties("unit_string").Value = "gn"
        End If
        If Chn.Properties("unit_string").Value = "Deg/Sec" Then
            Chn.Properties("unit_string").Value = "°/s"
        End If
        Chn.Properties("wf_increment").Value = 0.0001
        If Chn.Name <> "10VEHC0000__VO0" And Chn.Name <> "10VEHC0000__VOA" Then
          Call ChnOffset(Chn,Chn,100,"mean value offset")
        End If
      End If
    Next
    
    Call oGroup.Channels.Remove("Time axis")
    Call ChnLinGenImp("Time axis",6001,-0.1,0.0001,"")
    Call oGroup.Channels("Time axis").Properties.Add("Test ID", iihsTestId, DataTypeString)
    Call oGroup.Channels("Time axis").Properties.Add("Test Type", testDescription, DataTypeString) 
    oGroup.Name = iihsTestId & "_Raw_Data"
  End Sub
  
  Sub CleanOldFrontalDataScript
    Dim group : Set Group = Data.Root.ChannelGroups(1)
    If group.Channels.Exists("11VEHC0000__VO0") Then group.Channels("11VEHC0000__VO0").Name = "10VEHC0000__VO0"
    If group.Channels.Exists("11VEHC0000__ACX_") Then group.Channels("11VEHC0000__ACX_").Name = "10VEHC0000__ACX_"
    If group.Channels.Exists("11VEHC0000__ACY_") Then group.Channels("11VEHC0000__ACY_").Name = "10VEHC0000__ACY_"
    If group.Channels.Exists("11VEHC0000__ACZ_") Then group.Channels("11VEHC0000__ACZ_").Name = "10VEHC0000__ACZ_"
  End Sub

  Sub AskToConvertEvents(testType)
    If msgbox("It looks like kinematic events were added to this test using an outdated version of the dialog box. Would you like to import them into the latest version?", vbYesNo) = vbYes Then
      If testType = 1 Then Call ConvertModEvents
      If testType = 5 Then Call ConvertSmallEvents
    End If
  End Sub

  Sub ConvertModEvents
    If Not data.Root.ChannelGroups.Exists("Film_Analysis") Then Exit Sub
    Dim oGroup : Set oGroup = data.Root.ChannelGroups("Film_Analysis")
    Dim hardCodedEvents : hardCodedEvents = Array("Belt_tensioner_deploy_time", "Dr_airbag_time", "RF_airbag_time", "Face_time", "Airbag_full_time", "Face_loading_time", "SideCurtainDeployment")
    Dim hardCodedEventsId : hardCodedEventsId = Array(1, 2, 82, 5, 6, 7, 3)

    Call ConvertHardCodedEvents(oGroup, hardCodedEvents, hardCodedEventsId)
    Call ConvertAdnlEvents(oGroup)
  End Sub

  Sub ConvertSmallEvents
    If Not data.Root.ChannelGroups.Exists("Analysis") Then Exit Sub
    Dim oGroup : Set oGroup = data.Root.ChannelGroups("Analysis")
    Dim hardCodedEvents : hardCodedEvents = Array("Dr_airbag_time", "curtain_airbag_time", "Torso_airbag_time", "Face_time", "Airbag_full_time", "Face_loading_time")
    Dim hardCodedEventsId : hardCodedEventsId = Array(2, 3, 4, 5, 6, 7)

    Call ConvertHardCodedEvents(oGroup, hardCodedEvents, hardCodedEventsId)
    Call ConvertAdnlEvents(oGroup)
  End Sub

  Sub ConvertHardCodedEvents(oGroup, propNameArray, propValueArray)
    Dim i
    For i = 0 To Ubound(propNameArray)
      If oGroup.Properties.Exists(propNameArray(i)) Then
        If IsNumeric(oGroup.Properties(propNameArray(i)).Value) Then
          Call AddKinEvent(oGroup, propValueArray(i), oGroup.Properties(propNameArray(i)).Value) 
        End If
      End If
    Next

  End Sub
  
  Sub ConvertAdnlEvents(oGroup)
    Dim propertyName,idPropName , i, timePropName
    For i = 1 To 6
      propertyName = "Adnl_Evnt_" & i
      idPropName = propertyName & "ID"
      timePropName = propertyName & "tm"
      If oGroup.Properties.Exists(idPropName) And oGroup.Properties.Exists(timePropName) Then
        If IsNumeric(oGroup.Properties(idPropName).Value) Then
          Call AddKinEvent(oGroup, oGroup.Properties(idPropName).Value, oGroup.Properties(timePropName).Value) 
        End If
      End If
    Next

  End Sub

  Sub AddKinEvent(group, eventId, eventTime)
    Dim EventIdChn : EventIdChn = "11EventId"
    Dim EventTimeChn : EventTimeChn = "11EventTime"

    If Not group.Channels.Exists(EventIdChn) Or Not group.Channels.Exists(EventTimeChn) Then Exit Sub
    If group.Channels(EventIdChn).Size <> group.Channels(EventTimeChn).Size Then Exit Sub

    Dim insertLocation : insertLocation = group.Channels(EventIdChn).Size + 1
    group.Channels(EventIdChn).Values(insertLocation) = eventId
    group.Channels(EventTimeChn).Values(insertLocation) = cint(eventTime)
  End Sub