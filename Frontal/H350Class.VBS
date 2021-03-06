'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 03/23/2017 18:19:25
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\HeadChannelsClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\NeckChannelsClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\ChestChannelsClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\LegChannelsClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\Frontal\CorridorsChannelsClass.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\ConfigFileFactory.VBS")
Call ScriptInclude(APPPATH & "crashworthiness\CrashTools.VBS")

Class H350
  Private iBodyRegionList, iTestTypeId, iWorkbookPath, oConfigFileFactory, iOccupantLocation
  Public Property Get ChannelDictionary
    Set ChannelDictionary = iChannelDictionary
  End Property

  Public Default Function Init(occupantLocation, testTypeId) 'Use vraTestTypes 1="Moderate Overlap" 5="Small Overlap" 11="Right Side Small Overlap"
    iTestTypeId = testTypeId
    iOccupantLocation = occupantLocation
    Set iBodyRegionList = CreateObject("System.Collections.ArrayList")
    Set oConfigFileFactory = GetConfigFileFactory(testTypeId)
    iBodyRegionList.Add((New HeadChannels)(occupantLocation, oConfigFileFactory))
    iBodyRegionList.Add((New NeckChannels)(occupantLocation, oConfigFileFactory))
    iBodyRegionList.Add((New ChestChannels)(occupantLocation, oConfigFileFactory))
    iBodyRegionList.Add((New LegChannels)(occupantLocation, "leftLegGroupName", oConfigFileFactory))
    iBodyRegionList.Add((New LegChannels)(occupantLocation, "rightLegGroupName", oConfigFileFactory))
    iBodyRegionList.Add((New CorridorsChannels)(occupantLocation, oConfigFileFactory))
    GetConfigProperties
    Set Init = Me
  End Function
 
  Private Sub GetConfigProperties
    iWorkbookPath = APPPATH & oConfigFileFactory.GetConfigValue("testSettings", "workbookName", "path")
  End Sub
  
  Public Sub FilterChannels
    Dim region
    LogError("Filtering for occupant: " & iOccupantLocation & " started...")
    For Each region In iBodyRegionList
      region.FilterChannels
    Next
    LogError("Filtering for occupant: " & iOccupantLocation & " finished...")
  End Sub
  
  Public Sub CalculateInjuryCriteria
    Dim region
    LogError("Injury calculations for occupant: " & iOccupantLocation & " started...")
    For Each region In iBodyRegionList
      region.CalculateInjuryCriteria
    Next
    LogError("Injury calculations for occupant: " & iOccupantLocation & " finished...")
  End Sub
  
  Public Sub ExportProperties(workbook)
    Dim sheetName : sheetName = oConfigFileFactory.GetConfigValue("testSettings", "exportSheets", "occupant")
    sheetName = iOccupantLocation & sheetName
    Dim sheet : Set sheet = workbook.Sheets(sheetName)
    sheet.Select
    Call TransferPropToExcel(sheet, "iihsTestId" , data.Root, "Test_ID", "", "")
    Call TransferPropToExcel(sheet, "vehicleDescription" , data.Root, "Vehicle_Name", "", "")
    Call TransferPropToExcel(sheet, "testDescription" , data.Root, "Description", "", "")
    Call TransferPropToExcel(sheet, "impactSpeed" , data.Root, "Impact_velocity", 1, 1)    
    Call TransferPropToExcel(sheet, "testDate" , data.Root, "Test_Date", "", "")
    sheet.Range("createDate").Value = date()
    
    Dim region
    For Each region In iBodyRegionList
      region.ExportProperties(sheet)
    Next 
    
    Call workbook.Application.Run("Highlight")

    Set sheet = Nothing
  End Sub
End Class