'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/30/2017 10:41:54
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")


Class TestProfileBase
  Private iCrashPath, iDiademPath, iExcelPath, iDasPath, iConfigfactory, iObjectList, iIihsTestId, iVersion
  
  Public Property Get CrashPath
    CrashPath = iCrashPath
  End Property
  
  Public Property Get DiademPath
    DiademPath = iDiademPath
  End Property
  
  Public Property Get ExcelPath
    ExcelPath = iExcelPath
  End Property
  
  Public Property Get DasPath
    DasPath = iDasPath
  End Property
  
  Public Property Get IihsTestId
    IihsTestId = iIihsTestId
  End Property
  
  Public Property Get ObjectList
    Set ObjectList = iObjectList
  End Property
  
  Public Default Function Init(configFactory, iihsTestId)
    Set iConfigFactory = configFactory
    iIihsTestId = iihsTestId
    Set iObjectList = CreateObject("System.Collections.ArrayList")
    Call GetConfigValues
    Call Data.Root.Properties.Add("version", iVersion)
    Call Data.Root.Properties.Add("Data_Path", iCrashPath & "\" & iIihsTestId, DataTypeString)
    Set Init = Me
  End Function
  
  Public Sub FilterTest
    Dim item
    For Each item In iObjectList
      item.FilterChannels
    Next
  End Sub
  
  Public Sub ProcessTest
    Dim item
    For Each item In iObjectList
      item.CalculateInjuryCriteria
    Next
  End Sub
  
  Private Sub GetConfigValues
    iCrashPath = iConfigFactory.GetConfigValue("testSettings", "saveDirectory", "path")
    iDiademPath = iCrashPath & "\" & iIihsTestId & iConfigFactory.GetConfigValue("testSettings", "saveDirectory", "diademPath")
    iExcelPath = iCrashPath & "\" & iIihsTestId & iConfigFactory.GetConfigValue("testSettings", "saveDirectory", "excelPath")
    iVersion = iConfigFactory.GetConfigValue("testSettings", "version", "value")
    iDasPath = iConfigFactory.GetConfigValue("testSettings", "dasDirectory", "path") & "\" & iIihsTestId
  End Sub
  
  Public Sub CopyDas
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
      Call fso.CopyFolder(iDasPath, iCrashPath & "\" & iIihsTestId & iConfigFactory.GetConfigValue("testSettings", "dasDirectory", "path2"), True)
    If Err.number <> 0 Then LogError("DAS copy failed. Message: " & Replace(Err.Description, vbCrLf, ""))
    On Error Goto 0 
  End Sub
    
  Public Sub SaveTest    
    
    Call DataFileSave(iDiademPath & "\" & iIihsTestId & ".TDMS","TDMS")
  End Sub
  
End Class