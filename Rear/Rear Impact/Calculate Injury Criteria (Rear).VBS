'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 04/27/2010 16:12:56
'-- Author: 
'-- Comment: Injury Criteria
'-------------------------------------------------------------------------------

'Calculate the Head Resultant and HIC36
Call ChnXYZAbsValue("[2]/S1HEAD0000__ACXA","[2]/S1HEAD0000__ACYA","[2]/S1HEAD0000__ACZA","/HeadResultant") '... X,Y,Z,E 
Call ChnHICCalc("","[3]/HeadResultant",1,0,0,0,10,0) '... XW,Y,HIC36,HIC15,HICVar,HICUserDef,HICUser,HICd 

'Create channel properties for the HIC36, HIC36_T1, and HIC36_T2
Call Data.Root.ChannelGroups(3).Channels("HeadResultant").Properties.Add("HIC36", HICRes(1,2), DataTypeFloat64)
Call Data.Root.ChannelGroups(3).Channels("HeadResultant").Properties.Add("HIC36_T1", HICRes(2,2), DataTypeFloat64)
Call Data.Root.ChannelGroups(3).Channels("HeadResultant").Properties.Add("HIC36_T2", HICRes(3,2), DataTypeFloat64)

'Average the right and left T1 acceleration
Call ChnAverage("'[2]/S1THSP01LE__ACXD', '[2]/S1THSP01RI__ACXD'","/AverageT1XG") '... ChnNoStr,E 

'Calculate the NIC
Call ChnUnitConvert("[3]/AverageT1XG","[3]/AverageT1XSI","m/s^2") '... Y,E,TargetUnit  
Call ChnNICRearCalc("","[2]/S1HEAD0000__ACXDSI","[3]/AverageT1XSI","[3]/NIC") '... XW,Y,Y1,E 
Data.Root.ChannelGroups(3).Channels("NIC").Name = "NIC"

'Calculate the Occipital Y Moment
Call Calculate("Ch(""[3]/MOCy"") = (Ch(""[2]/S1NECKUP00__MOYB"")-Ch(""[2]/S1NECKUP00__FOXB"")*.01778)",NULL,NULL,"") '... CalculateFormula,CalculateSymbols,CalculateValues,CalculateTargetUnit)

'Calculate the NIJ
Call ChnNijCalc("","[2]/S1NECKUP00__FOZB","[3]/MOCy","/NIJ_NTE",6806,-135,"NTE") '... XW,Y,Y1,E,NijFzc,NijMyc,NijType 
Call ChnNijCalc("","[2]/S1NECKUP00__FOZB","[3]/MOCy","/NIJ_NTF",6806,310,"NTF") '... XW,Y,Y1,E,NijFzc,NijMyc,NijType
Call ChnNijCalc("","[2]/S1NECKUP00__FOZB","[3]/MOCy","/NIJ_NCE",-6160,-135,"NCE") '... XW,Y,Y1,E,NijFzc,NijMyc,NijType 
Call ChnNijCalc("","[2]/S1NECKUP00__FOZB","[3]/MOCy","/NIJ_NCF",-6160,310,"NCF") '... XW,Y,Y1,E,NijFzc,NijMyc,NijType 

'Calculate the NKM
Call ChnNijCalc("","[2]/S1NECKUP00__FOXB","[3]/MOCy","/Nkm_NEA",845,-47.5,"NTE") '... XW,Y,Y1,E,NijFzc,NijMyc,NijType 
Call ChnNijCalc("","[2]/S1NECKUP00__FOXB","[3]/MOCy","/Nkm_NFA",845,88.1,"NTF") '... XW,Y,Y1,E,NijFzc,NijMyc,NijType
Call ChnNijCalc("","[2]/S1NECKUP00__FOXB","[3]/MOCy","/Nkm_NEP",-845,-47.5,"NCE") '... XW,Y,Y1,E,NijFzc,NijMyc,NijType 
Call ChnNijCalc("","[2]/S1NECKUP00__FOXB","[3]/MOCy","/Nkm_NFP",-845,88.1,"NCF") '... XW,Y,Y1,E,NijFzc,NijMyc,NijType 

'Calculate the Head Rebound Velocity
Call ChnIntegrate("[1]/Time axis","[2]/S1HEAD0000__ACXC","/Integrated") '... XW,Y,E 
Call ChnIntegrate("[1]/Time axis","[2]/S0VEHC0300__ACXD","/Integrated1") '... XW,Y,E 
Call ChnToWfChn("[1]/Time axis","'[3]/Integrated' - '[3]/Integrated1'",0) '... X,ChnNoStr,XChnDelete 
Call ChnLinScale("[3]/Integrated","/Integrated",9.8,0) '... Y,E,ChnScaleFactor,ChnScaleOffset 
Call ChnLinScale("[3]/Integrated1","/Integrated1",9.8,0) '... Y,E,ChnScaleFactor,ChnScaleOffset 
Call ChnSub("[3]/Integrated","[3]/Integrated1","/Subtracted") '... Y,Y1,E 
Call Data.Root.ChannelGroups(3).Channels.Remove("Integrated")
Call Data.Root.ChannelGroups(3).Channels.Remove("Integrated1")
Data.Root.ChannelGroups(3).Channels("Subtracted").Name = "Head_Rebound_Velocity"
Call Data.Move(Data.Root.ChannelGroups(3).Channels("Head_Rebound_Velocity"),Data.Root.ChannelGroups(3).Channels,4)

'Calculate T1 - Head Acceleration
Call ChnSub("[3]/AverageT1XG","[2]/S1HEAD0000__ACXA","/T1XG_HeadXG") '... Y,Y1,E 
Call Data.Move(Data.Root.ChannelGroups(3).Channels("T1XG_HeadXG"),Data.Root.ChannelGroups(3).Channels,3)

