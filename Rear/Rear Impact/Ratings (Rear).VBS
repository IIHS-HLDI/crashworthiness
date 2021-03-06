'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 09/10/2010 16:00:49
'-- Author: 
'-- Comment: Load Neck Force Corridors and determine the ratings
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

Call DataFileLoad(CurrentScriptPath & "Neck Force Corridors.tdms","TDMS","Load") '... DataFilename,FileImportFilter,ImportAction 

'Calculate the Neck Force Rating
Dim Max_Fx
Dim Max_Fz
Dim TempValue1
Dim TempValue2
Dim TempValue3
Dim Neck_Force_Rating

Max_Fx = Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOXA").Properties("Max_Fx").Value
Max_Fz = Data.Root.ChannelGroups(2).Channels("S1NECKUP00__FOZA").Properties("Tension").Value

'Add max shear and tension values as properties
Call Data.Root.Properties.Add("Maximum Shear Force(N)", Max_Fx, datatypefloat64)
Call Data.Root.Properties.Add("Maximum Tension(N)", Max_Fz, datatypefloat64)

'Calculate the neck rating
TempValue1 = (Max_Fx/315)^2
TempValue2 = ((Max_Fz - 234)/1131)^2
TempValue3 = TempValue1 + TempValue2


If TempValue3 < 0.2025 then
  Neck_Force_Rating = "LOW"
  Else 
    If TempValue3 > 0.680625 then
    Neck_Force_Rating = "HIGH"
    Else
    Neck_Force_Rating = "MODERATE"
    End IF
End IF

'Add Neck Force Rating to TDMS properties
Call Data.Root.Properties.Add("Neck Forces Rating", Neck_Force_Rating, Datatypestring)

'Display the Seat Design Characteristics
Dim Max_T1
Dim HCT
Dim Max_T1_Criteria
Dim HCT_Criteria

'Determine peak acceleration, deltaV, acceleration at t0, and endtime from sled acceleration #3
Max_T1 = Data.Root.ChannelGroups(3).Channels("AverageT1XG").Properties("Max_T1Xg").Value

HCT = (Data.Root.ChannelGroups(3).Channels("HCT").Properties("HCT_begin").Value)*1000

'Determine if the T1 and Head contact time are below the reference values
If Max_T1 <= 9.5 then
  Max_T1_Criteria = "PASS"
  Else 
  Max_T1_Criteria = "FAIL"
End IF

If HCT <= 70 then
  HCT_Criteria = "PASS"
  Else 
  HCT_Criteria = "FAIL"
End IF

'Add Seat Design Characteristic Criteria to TDMS properties
Dim Dynamic_Rating

Call Data.Root.Properties.Add("Maximum T1 Acceleration(g)", Max_T1, Datatypefloat64)
Call Data.Root.Properties.Add("Maximum T1 Criteria", Max_T1_Criteria, Datatypestring)
Call Data.Root.Properties.Add("Head Contact Time(ms)", HCT, Datatypefloat64)
Call Data.Root.Properties.Add("HCT Criteria", HCT_Criteria, Datatypestring)

'Determine Dynamic Rating
IF ((Max_T1 <= 9.5 or HCT <= 70) and Neck_Force_Rating = "LOW") then
  Dynamic_Rating = "Good"
  ElseIf ((Max_T1 <= 9.5 or HCT <= 70) and Neck_Force_Rating = "MODERATE") then
  Dynamic_Rating = "Acceptable"
  ElseIF ((Max_T1 <= 9.5 or HCT <= 70) and Neck_Force_Rating = "HIGH") then
  Dynamic_Rating = "Marginal"
  ElseIf ((Max_T1 > 9.5 and HCT > 70) and Neck_Force_Rating = "LOW") then
  Dynamic_Rating = "Acceptable" 
  ElseIf ((Max_T1 > 9.5 and HCT > 70) and Neck_Force_Rating = "MODERATE") then
  Dynamic_Rating = "Marginal"
  Else
  Dynamic_Rating = "Poor"
End IF

Call Data.Root.Properties.Add("Dynamic Rating", Dynamic_Rating, Datatypestring)

'Determine Overall Rating

Dim Overall_Rating
Dim Static_Rating
Dim Rated_Height


IF (Static_Rating= "Acceptable" and Rated_Height <=60 and Dynamic_Rating= "Good") then
    Overall_Rating= "Good"
    ELSEIF(Static_Rating= "Acceptable" and Dynamic_Rating= "Good") then
    Overall_Rating= "Acceptable"
    Else
    Overall_Rating=Dynamic_Rating
End IF
Call Data.Root.Properties.Add("Overall Rating", Overall_Rating, Datatypestring)

'Delete the Tension channel
'Call Data.Root.ChannelGroups(6).Channels.Remove("Tension")