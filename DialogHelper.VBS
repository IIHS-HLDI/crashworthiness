'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 08/07/2017 09:32:25
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
  Class EventsData
    Private iEvents, channelNameArray, channelDataTypeArray, iChannels, group
    
    Public Property Get EventChannels
      Set EventChannels = iChannels
    End Property
    
    Public Property Get Events
      Set Events = iEvents
    End Property
    
    Public Default Function Init(oGroup, occupantPosition)
      Set group = ogroup
      channelNameArray = Array(occupantPosition & "EventId", occupantPosition & "EventTime")
      channelDataTypeArray = Array(DataTypeFloat64, DataTypeFloat64)
      Set iChannels = CreateObject("System.Collections.ArrayList")
      SetupChannels
      Set iEvents = CreateObject("System.Collections.ArrayList")
      GetEvents
      Set Init = Me
    End Function
    
    Private Sub SetupChannels
      Dim name, i
      i = 0
      For Each name In channelNameArray
        If Not group.Channels.Exists(name) Then
          Call group.Channels.Add(name, channelDataTypeArray(i))
        End If
        Call iChannels.Add(group.Channels(name))
        i = i + 1
      Next 
    End Sub
    
    Public Sub SortEvents(table)
      Call ChnMultipleSortExt(iChannels(1), iChannels(0), "Up", "CaseInsensitive", True)
      table.Refresh
    End Sub
    
    Public Sub LaunchEventManager
      Call ExtProgram("iexplore.exe", "http://applications1/VehicleResearchApplication/IIHS.VehicleResearch.App.application?ShowEventManager&0")
    End Sub
    
    Public Sub AddEvent(ByRef table)
      table.RowCount = table.RowCount + 1
      iChannels(0).Values(table.RowCount) = -1
      iChannels(1).Values(table.RowCount) = 0 
    End Sub
    
    Public Sub RemoveEvent(table)
      If Not table.ActiveCellRow = -1 Then
        Call iChannels(0).RemoveValues(table.ActiveCellRow, 1)
        Call iChannels(1).RemoveValues(table.ActiveCellRow, 1) 
        table.RunInitialize
        table.Refresh
      End If
    End Sub
    
    Public Sub GetEvents
      dim odbConn
      dim RS
      dim cmConn
      iEvents.Clear
      Call iEvents.Add(Array("", -1))
      Set oDBConn = CreateObject("ADODB.Connection")
      
      oDBConn.ConnectionString = "Provider=SQLOLEDB.1;Initial Catalog=VRCData;Data Source=vrcvapp;Integrated Security=SSPI;Persist Security Info=False"
      oDBConn.Open
      
      Set RS = oDBConn.Execute("Select * From kin_Event ORDER BY EventName")
      
      If not RS.EOF Then
        RS.MoveFirst
        Do While not RS.EOF
          If RS.Fields("VRATestTypeId") = 1 Then
            Call iEvents.Add (Array(RS.Fields("EventName").Value, RS.Fields("Id").Value))
          End If
        RS.MoveNext
        Loop 
      End If
      RS.Close
      
      Set oDBConn = Nothing
      Set RS = Nothing
    End Sub
    
  End Class
  
  Sub GetTableValues(eventData, table, Row, Col, ByRef Cell, IsInputCell)
  Dim chns : Set chns = eventData.EventChannels
  If Row = 0 Then              ' set title of table
    Select Case Col
    Case 0, 3
      Cell.Text = ""           ' row with line numbers doesn't contain title
    Case Else
      Cell.Text = chns(Col - 1).Name ' display channel name
    End Select 
  Else
    Select Case Col
    Case 0
      Cell.Text = Row          ' display line number
    Case 1
      Call SetupLookup(eventData, Cell)
      If Not IsNull(chns(Col - 1).Values(Row)) Then Cell.Value = chns(Col - 1).Values(Row)
    Case 3
      Cell.Text = "..."
    Case Else
      Cell.Text = Str(chns(Col - 1).Values(Row))  ' display channel value
    End Select 
  End If
End Sub

Sub SetupLookup(eventData, ByRef cell)
  Dim item
  cell.Items.RemoveAll
  
  For Each item In eventData.Events
    
    Call cell.Items.Add(item(0), item(1))
  Next
End Sub

Sub SetTableValues(chns, table, Row, Col, ByRef Cell)
  If Row > 0 And Col > 0 Then
    Select Case Col
      Case 1
        chns(Col - 1).Values(Row) = Val(Eval(Cell.Value))  ' entered value included in channel
      Case 3
      Case Else
        chns(Col - 1).Values(Row) = Val(Eval(Cell.Text))  ' entered value included in channel
    End Select
  End If
End Sub