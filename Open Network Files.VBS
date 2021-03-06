'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 07/06/2010 09:15:46
'-- Author: 
'-- Comment: 
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

'This VBS is called by the 'User Form initializ.VBS'
 
Dim TESTID,testyr_34, test_year, Archived_path, CE_path, network_available, Test_Folder_path, Test_Folder_path2, IIHS_path

Call Data.Root.Clear() 
IIHS_path = "\\IIHS\IIHSDrive\VRC\Shared\"  
TESTID = UCase(T1)

Call Get_Test_Year 'This also identifies test type and sets DataReadPath to the correct location

CE_path = DataReadPath
Archived_path = CE_path & "Archived\"



'Confirm connection to IIHS network
Call Check_Network_Connection
IF network_available = FALSE THEN 'Network not available.  Prompt user to open file locally.
  msgbox(IIHS_path & " network location not available.")
  Dim MyVar
  MyVar = msgbox(IIHS_path & " network location not available.  Would you like to manually browse the local drive for DIAdem files?",4)
  IF MyVar = vbYes THEN   'Network is available
   
    Call FileNameGet("NAVIGATOR","FileRead",,"TDMS")
     
    Call DataFileLoad(FileDlgFileName)
    Call Collapse_Portal
    
    Call FileNameGet("VIEW","FileRead",,"TDV")
    Call View.LoadLayout(FileDlgFileName)
    Call WndShow("VIEW", "MAXIMIZE")    
  ELSE
  
    Call Show_UserForm
  END IF
ELSE 'Network connection is available
 
  'Check for Test Folder in different locations
  'First check the main folder
  Test_Folder_Path = CE_path & TESTID & "\DATA\DIAdem\"
  Test_Folder_Path2 = CE_path & TESTID
  Call Check_Test_folder
  'If folder not found try the Archived\CE folder followed by the Archived\R&D folder
  IF network_available = FALSE THEN
    Test_Folder_Path = Archived_path & test_year & "\CE\" & TESTID & "\DATA\DIAdem\"
    Test_Folder_Path2 = Archived_path & test_year & "\CE\" & TESTID 
    Call Check_Test_folder
    IF network_available = FALSE THEN
      Test_Folder_Path = Archived_path & test_year & "\R&D\" & TESTID & "\DATA\DIAdem\"
      Test_Folder_Path2 = Archived_path & test_year & "\R&D\" & TESTID 
      Call Check_Test_folder 
       IF network_available = FALSE THEN
       Test_Folder_Path = Archived_path & test_year & "\IAG\" & TESTID & "\DATA\DIAdem\"
       Test_Folder_Path2 = Archived_path & test_year & "\IAG\" & TESTID 
       Call Check_Test_folder 
      END IF
    END IF
  END IF
  
  IF network_available = FALSE THEN  'The test folder has not been found in the DATA\DIAdem folder. . .have the user browse for correct folder
    msgbox(TESTID & " DIAdem folder not found.  Browse to the folder containing the DIAdem files (TDMS and TDV).")
    Call FileNameGet("NAVIGATOR","FileRead",CE_path,"TDMS")
    IF Dlgstate = "IDOk" THEN  'User selected a TDMS file and clicked OK button. . .next open TDV file
      Call DataFileLoad(FileDlgFileName)
      
    End IF
    
    Call FileNameGet("VIEW","FileRead",CE_path,"TDV") 
    IF Dlgstate = "IDOk" THEN
      Call WndShow("VIEW", "MAXIMIZE")
      Call View.LoadLayout(FileDlgFileName)
    END IF
       
    Call Collapse_Portal
  
    
  ELSE 
    
    Call FileNameGet("NAVIGATOR","FileRead",Test_Folder_path,"TDMS")
    IF Dlgstate = "IDOk" THEN  'User selected a TDMS file and clicked OK button. . .next open TDV file
      Call DataFileLoad(FileDlgFileName)
      Call Data.Root.Properties.Add("Data_Path", Test_Folder_path2, DataTypeString)
      
    END IF
    
    Call FileNameGet("VIEW","FileRead",Test_Folder_path,"TDV")
    IF Dlgstate = "IDOk" THEN
      Call WndShow("VIEW", "MAXIMIZE")
      Call View.LoadLayout(FileDlgFileName)
    END IF
    
    If Left(TestID,3) = "SER" then
      Call Get_Report_View
    End If
    
    Call Collapse_Portal
   
    
 END IF
  
 
END IF



'------------------------------------------
Sub Collapse_Portal

Call UIAutoRefreshSet(True) 
    Call Portal.Structure.Collapse(Data.Root) 
    Call UIAutoRefreshSet(False) 
End Sub
'------------------------------------------
Sub Check_Test_folder
  dim filesys, newfolder
  set filesys=CreateObject("Scripting.FileSystemObject")
  IF  NOT filesys.FolderExists(Test_Folder_Path) THEN
    network_available = FALSE
  ELSE
    network_available = TRUE
  END If
  
  'msgbox(test_folder_path & " " & network_available)

End Sub
'------------------------------------------
'------------------------------------------
Sub Check_Network_Connection
 
dim filesys, newfolder
set filesys=CreateObject("Scripting.FileSystemObject")
  IF  NOT filesys.FolderExists(IIHS_path) THEN
    network_available = FALSE
    
  ELSE
    network_available = TRUE
    
  END If

End Sub
'------------------------------------------
Sub Get_Test_Year

'Determine the year the test was conducted
    IF LEFT(TESTID,2) = "CF" THEN
      testyr_34 = RIGHT(LEFT(TESTID,4),2)
      DataReadPath = IIHS_path & "Crashworthiness Program\Frontal Crash Tests\" 
    ELSEIF LEFT(TESTID,2) = "CN" THEN
      testyr_34 = RIGHT(LEFT(TESTID,4),2)
      DataReadPath = IIHS_path & "Crashworthiness Program\Frontal Crash Tests\" 
    ELSEIF LEFT(TESTID,2) = "CP" THEN
      testyr_34 = RIGHT(LEFT(TESTID,4),2)
      DataReadPath = IIHS_path & "Crashworthiness Program\Frontal Crash Tests\" 
    ELSEIF LEFT(TESTID,3) = "CEN" THEN
      testyr_34 = RIGHT(LEFT(TESTID,5),2)
      DataReadPath = IIHS_path & "Crashworthiness Program\Frontal Crash Tests\" 
    ELSEIF LEFT(TESTID,3) = "CEF" THEN
      testyr_34 = RIGHT(LEFT(TESTID,5),2)
      DataReadPath = IIHS_path & "Crashworthiness Program\Frontal Crash Tests\" 
    ELSEIF LEFT(TESTID,3) = "CEP" THEN
      testyr_34 = RIGHT(LEFT(TESTID,5),2)
      DataReadPath = IIHS_path & "Crashworthiness Program\Frontal Crash Tests\" 
    ELSEIF LEFT(TESTID,2) = "CS" THEN
      testyr_34 = RIGHT(LEFT(TESTID,4),2)
      DataReadPath = IIHS_path & "Crashworthiness Program\Side-Impact Crash Tests\" 
    ELSEIF LEFT(TESTID,3) = "CES" THEN
        testyr_34 = RIGHT(LEFT(TESTID,5),2)
        DataReadPath = IIHS_path & "Crashworthiness Program\Side-Impact Crash Tests\" 
    ELSEIF LEFT(TESTID,2) = "SR" THEN
      testyr_34 = RIGHT(LEFT(TESTID,4),2)
      DataReadPath = IIHS_path & "Crashworthiness Program\Rear Impact Sled Tests\" 
    ELSEIF LEFT(TESTID,3) = "SER" THEN
        testyr_34 = RIGHT(LEFT(TESTID,5),2)  
        DataReadPath = IIHS_path & "Crashworthiness Program\Rear Impact Sled Tests\" 
    END IF

    IF testyr_34 < 94 THEN   'This will work for tests from 2000 to 2093
      test_year = 2000 + testyr_34
    ELSE
      test_year = 1900 + testyr_34
    END IF
    
    
    

End Sub
'------------------------------------------

Sub Get_Report_View

  dim filesys
  dim reportFolder
  dim T2
  
  T2 = Data.Root.Properties("Vehicle_Name").Value
  reportFolder = "\\IIHS\IIHSDrive\VRC\Shared\Crashworthiness Program\Rear Impact Sled Tests\"
  
  Set filesys=CreateObject("Scripting.FileSystemObject")
   If (filesys.FileExists(reportFolder & T1 & "\DATA\DIAdem\" & T1 & " - " & T2 & ".TDR")) Then
       IF (FileNameGet("REPORT","FileRead",reportFolder & T1 & "\DATA\DIAdem\","TDR") = "IDOk") then
        Call PicLoad(FileDlgFileName)
        Call PicUpdate(0)
       END IF
  END IF
  
 End Sub