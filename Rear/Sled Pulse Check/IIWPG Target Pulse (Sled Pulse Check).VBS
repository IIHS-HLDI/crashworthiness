'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 09/10/2010 16:00:49
'-- Author: 
'-- Comment: Add IIWPG Target Pulse
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.


'-------------------------------------------------------------------------------
Call DataFileLoad(CurrentScriptPath & "Target Pulse.tdms","TDMS","Load") '... DataFilename,FileImportFilter,ImportAction 

'Call DataFileLoad(IIWPGPulse(0)&"EuroNCAP corridors1.tdms","TDMS","Load") '... DataFilename,FileImportFilter,ImportAction 