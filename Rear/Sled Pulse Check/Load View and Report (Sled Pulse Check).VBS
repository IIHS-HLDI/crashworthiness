'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 09/13/2010 08:57:39
'-- Author: 
'-- Comment: Load Sled Pulse Check View and Report
'-------------------------------------------------------------------------------
Option Explicit  'Forces the explicit declaration of all the variables in a script.


'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Call View.NewLayout()
Call View.LoadLayout(CurrentScriptPath & "Pulse Check.TDV")

'-------------------------------------------------------------------------------
Call PicLoad(CurrentScriptPath & "Sled Pulse Check.TDR") '... PicFile 
Call PicUpdate(0)                       '... PicDoubleBuffer 