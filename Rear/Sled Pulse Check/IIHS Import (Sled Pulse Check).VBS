'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/23/2010 13:28:34
'-- Author: 
'-- Comment: Sled Pulse Import
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.

dim chnarray(0)
chnarray(0) = "Time axis"
View.AutoRefresh = TRUE
Call DataFileLoad(FileDlgName,"DD8x","")
'Call DataFileLoad(FileDlgName,"DAT","") '... DataFilename,FileImportFilter,ImportAction 


'Call script to change numeric channels to waveforms, remove offset, and create filtered and calculated channel groups
Call Scriptinclude(CurrentScriptPath & "Change to Waveform and Remove offset (Sled Pulse Check).vbs") ', "RemoveOffset")

'Call script to filter data
Scriptinclude(CurrentScriptPath & "Filter Data (Sled Pulse Check).vbs")

'Call script to shift data so the peak sled acceleration occurs at 27 ms.
Scriptinclude(CurrentScriptPath & "Data shift (Sled Pulse Check).vbs")

'Call script to calculate delta V for the sled accelerometers
Scriptinclude(CurrentScriptPath & "Delta V calculations (Sled Pulse Check).vbs")

'Call script to display sled acceleration pulse specifications
Scriptinclude(CurrentScriptPath & "Pulse Specifications (Sled Pulse Check).vbs")

'Call script to import IIWPG target pulse
Scriptinclude(CurrentScriptPath & "IIWPG Target Pulse (Sled Pulse Check).vbs")

'Call script to load the view
Scriptinclude(CurrentScriptPath & "Load View and Report (Sled Pulse Check).vbs")

'Change the file name
Dim MyString, MyArray, Msg
MyString = Data.Root.Properties("name").Value
MyArray = Split(MyString, "_DIAdemHeader_Binary", -1, 1)
Data.Root.Properties("name").Value = MyArray(0)
Data.Root.ChannelGroups(1).Properties("name").Value = "Raw Data"



'Save files in test directory
Dim pulse_directory
Dim Name
Name = Data.Root.Properties("name").Value
pulse_directory =  "N:\VRC\Shared\Crashworthiness Program\Rear Impact Sled Tests\Pulse Check\" & Name

'Create folder for sled pulse check
Dim fso, f
Set fso = CreateObject("Scripting.FileSystemObject")

If FolderExist(pulse_directory) then 
  Call DataFileSave(pulse_directory & "\" & Name & ".tdms","TDMS") 'Save TDMS file
  Call View.SaveLayout(pulse_directory & "\" & Name & ".TDV") 'Save View file
  Call PicSaveAs(pulse_directory & "\" & Name & ".TDR","") 'Save Report file
  Else
  Set f = fso.CreateFolder(pulse_directory)
End IF

'Save TDMS file
Call DataFileSave(pulse_directory & "\" & Name & ".tdms","TDMS")

'Save View file in test directory
Call View.SaveLayout(pulse_directory & "\" & Name & ".TDV")

'Save Report file in test directory
Call PicSaveAs(pulse_directory & "\" & Name & ".TDR","") 

Call WndShow("Report", "Maximize")