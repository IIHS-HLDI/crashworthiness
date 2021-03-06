'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 12/03/2010 09:45:37
'-- Author: 
'-- Comment: Get sled test information
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

'Call Rear Impact Dialogue box
Call SUDDefLoad(CurrentScriptPath & "Rear Impact.sud")
Call SUDDlgShow("Dlg1")

IF DlgState = "IDOk" THEN 'OK button was pressed
  
  Call Data.Root.Properties.Add("ATD_ID", "0", DataTypeString) 'Create a custom property for the BioRID ID  

  'Select the correct BioRID  
  IF B1 = 0 THEN		
      Data.Root.Properties("ATD_ID").Value = "BioRID A"
    ELSE B1 = 1 		
      Data.Root.Properties("ATD_ID").Value = "BioRID B"
  END IF

  Call Data.Root.Properties.Add("ATD_Build_Level", "0", DataTypeString) 'Create a custom property for the BioRID build level 

  'Select the build level
  IF B2 = 0 THEN		
      Data.Root.Properties("ATD_Build_Level").Value = "IIg"
    ELSE B2 = 1 		
      Data.Root.Properties("ATD_Build_Level").Value = "Other"
  END IF

  Call Data.Root.Properties.Add("BioRID_Test_Number", "0", DataTypeString) 'Create a custom property for the BioRID test number
  
   Data.Root.Properties("BioRID_Test_Number").Value = T5

  Call Data.Root.Properties.Add("HPG_Protocol_Number", "0", DataTypeString) 'Create a custom property for the BioRID test number
  
   Data.Root.Properties("HPG_Protocol_Number").Value = T6

  Call Data.Root.Properties.Add("Rated_Backset", "0", DataTypeString) 'Create a custom property for the BioRID test number
  
   Data.Root.Properties("Rated_Backset").Value = T7

  Call Data.Root.Properties.Add("Rated_Height", "0", DataTypeString) 'Create a custom property for the BioRID test number
  
   Data.Root.Properties("Rated_Height").Value = T8
    
  Call Data.Root.Properties.Add("Static_Rating", "0", DataTypeString) 'Create a custom property for the BioRID test number
  
    Data.Root.Properties("Static_Rating").Value = T9
  

End IF

