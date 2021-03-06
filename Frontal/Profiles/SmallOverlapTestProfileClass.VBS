'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/27/2017 16:15:31
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\H350Class.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\FrontalVehicleChannelsClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ExcelExportClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\TestProfileBaseClass.VBS")

Class SmallOverlapTestProfile
  Private iDriver, iPassenger, iVehicle, iTestTypeId, iConfigFactory, iHasPassenger, iHasDriver, iBase
  
  Public Property Get Base
    Set Base = iBase
  End Property
  
  Public Default Function Init(iihsTestId)
    iTestTypeId = 6
    'This is an attempt at inheriting from a base class. iBase must be included on every profile.
    Set iConfigFactory = GetConfigFileFactory(iTestTypeId)
    Set iBase = (New TestProfileBase)(iConfigFactory, iihsTestId)
    
    If Data.Root.Properties("HasDriver").Value Then 
      iHasDriver = true
      Set iDriver = (New H350)("11", iTestTypeId)
      Call iBase.ObjectList.Add(iDriver)
    End If 
    If Data.Root.Properties("HasPassenger").Value Then 
      iHasPassenger = true
      Set iPassenger = (New H350)("13", iTestTypeId)
      Call iBase.ObjectList.Add(iPassenger)
    End If
    Set iVehicle = (New FrontalVehicleChannels)("1", iTestTypeId)
    Call iBase.ObjectList.Add(iVehicle)
    Set Init = Me
  End Function
  
  Public Sub ExportTest
    iBase.ProcessTest
    Dim path : path = APPPATH & iConfigFactory.GetConfigValue("testSettings", "workbookName", "path")
    Dim export : Set export = (New ExcelExport)(path)
    
    'Call ExportDialogData(export.Workbook)
    Dim item
    For Each item In iBase.ObjectList
      Call item.ExportProperties(export.Workbook)
    Next
    
    Call export.Workbook.Application.Run("Hide_MO_Chart")
    Call export.DeleteSheet("Film Analysis")
    Call export.Workbook.Application.Run("CleanupOccupants", iHasPassenger)
    Call export.SaveAs(iBase.ExcelPath, iBase.IihsTestId)
    Set export = Nothing
  End Sub
  
  Public Sub Run
    iBase.FilterTest
    iBase.ProcessTest
    iBase.SaveTest
    GetDiademViews
    ExportTest
    iBase.CopyDas
  End Sub
  
  Private Sub GetDiademViews
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sourcePath : sourcePath = APPPATH & "crashworthiness\Frontal\"
    Dim viewPath
    
    If iHasPassenger Then
      viewPath = sourcePath & iConfigFactory.GetConfigValue("testSettings", "diademViews", "passengerFiltered")
      Call fso.CopyFile(viewPath, iBase.DiademPath & "\" & iBase.IihsTestId & "  (RF Passenger).TDV")
      viewPath = sourcePath & iConfigFactory.GetConfigValue("testSettings", "diademViews", "passengerUnfiltered")
      Call fso.CopyFile(viewPath, iBase.DiademPath & "\" & iBase.IihsTestId & "_RF_unfiltered.TDV")
      Call View.LoadLayout(iBase.DiademPath & "\" & iBase.IihsTestId & "_RF_unfiltered.TDV")
    End If
    
    If iHasDriver Then 
      viewPath = sourcePath & iConfigFactory.GetConfigValue("testSettings", "diademViews", "driverFiltered")
      Call fso.CopyFile(viewPath, iBase.DiademPath & "\" & iBase.IihsTestId & ".TDV")
      viewPath = sourcePath & iConfigFactory.GetConfigValue("testSettings", "diademViews", "driverUnfiltered")
      Call fso.CopyFile(viewPath, iBase.DiademPath & "\" & iBase.IihsTestId & "_unfiltered.TDV")
      Call View.LoadLayout(iBase.DiademPath & "\" & iBase.IihsTestId & "_unfiltered.TDV")
    End If
    
  End Sub
  
  Private Sub ExportDialogData(workbook)
    Dim dialogData : Set dialogData = iConfigFactory.GetDialogExportList
    Dim sheetName : sheetName = iConfigFactory.GetConfigValue("testSettings", "exportSheets", "dialog")
    Dim group : group = iConfigFactory.GetConfigValue("testSettings", "dialogGroupName", "name")
    
    If Not Data.Root.ChannelGroups.Exists(Group) Then
      LogError("Failed to export dialog data. Group: " & group & " does not exist.")
      Exit Sub
    End If
    
    Dim sheet : Set sheet = workbook.Sheets(sheetName)
    sheet.Select
    Dim exportData
    
    For Each exportData In dialogData
      Call TransferPropToExcel(sheet, exportData.Cell , data.Root.ChannelGroups(group), exportData.PropertyName, "", "")
    Next
    
    Set dialogData = Nothing
    Set sheet = Nothing 
  End Sub
  
End Class