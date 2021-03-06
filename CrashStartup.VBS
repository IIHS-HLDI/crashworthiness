'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 08/03/2016 15:00:21
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(CurrentScriptPath & "CrashTools.VBS")
GlobalDim("CrashScriptPath")
CrashScriptPath = CurrentScriptPath
Dim TDAS_DataPath

TDAS_DataPath = "\\IIHS\IIHSDrive\VRC\Shared\Resources\Dummy Information\TDAS\"

Call Show_UserForm

IF DlgState = "IDOk" THEN 'OK button was pressed
  IF L2 = 1 THEN		'Process new crash data
    Call Data.Root.Clear()
    Call Process_Crash_Data
  ELSEIF L2 = 0 THEN		'Open existing TDMS file and VIEW layout
    Call Open_Existing_Files
  ELSEIF L2 = 2 THEN		'Open in Normal mode
    Call WndShow("NAVIGATOR", "MAXIMIZE")
  END IF
END IF

'------------------------------------------
Sub Process_Crash_Data
  Dim Test_Type
  If IsConnected Then   
    Select Case L1
      Case "0"  
        DataReadPath = TDAS_DataPath & "FRONT\DATA\" & T1
        Test_Type = "Frontal"
        Call ScriptInclude(CurrentScriptPath & Test_Type & "\IIHS import")
      Case "1"  
        DataReadPath = TDAS_DataPath & "SIDE\DATA\" & T1
        Test_Type = "Side"
        Call ScriptInclude(CurrentScriptPath & Test_Type & "\Side import")
      Case "2"  
        DataReadPath = TDAS_DataPath & "REAR\DATA\" & T1
        Test_Type = "Rear\Rear Impact"
        Call ScriptInclude(CurrentScriptPath & Test_Type & "\IIHS Import (Rear)","MainRearProcessing")
      Case "3"  
        DataReadPath = TDAS_DataPath & "REAR\DATA\" & T1  'This is reserved for the pulse check
        Test_Type = "Rear\Sled Pulse Check"
        If FileNameGet("Navigator","FileRead",T1 & "_DIAdemHeader_BINARY.dat", "DAT") = "IDCancel" Then
          msgbox("Data import cancelled.")
        Else
          Call ScriptInclude(CurrentScriptPath & Test_Type & "\IIHS import (Sled Pulse Check)")
        End If
      Case "4"  
        DataReadPath = TDAS_DataPath & "ROOF\DATA\" & T1
      Case Else
    End Select
  End If  
End Sub
'------------------------------------------
Sub Open_Existing_Files    
    'Save TDMS?
    If IsDataFileLoaded = true Then
      If msgbox("Save data file: " & Data.Root.Name, vbYesNo) = vbYes Then
        Call DataFileSave()
      End If
    End If
    
    'Save layout file?
    If IsViewLoaded = true Then
      If msgbox("Save layout file: " & View.FileName, vbYesNo) = vbYes THEN
        Call View.SaveLayout(View.FullPath)
      End If
    End If      
    ScriptInclude(CurrentScriptPath & "Open Network Files.VBS")   
End Sub
'------------------------------------------
Sub Show_UserForm
  Call SUDDlgShow("Dlg1", CurrentScriptPath & "CrashDialog.sud")
End Sub
