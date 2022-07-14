# LabChart_Heart_Rate Macro #

	Sub Heart_Rate ()
	
	Call Doc.OpenView ("Data Pad")
	
	' Begin DataPadColumnSetup
	Column = 1
	FunctionType = "Time"
	Channel = ##Heart Rate Channel##
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 2
	FunctionType = "Selection Start"
	Channel = ##Heart Rate Channel##
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 3
	FunctionType = "Selection End"
	Channel = ##Heart Rate Channel##
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 4
	FunctionType = "Selection Duration"
	Channel = ##Heart Rate Channel##
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 5
	FunctionType = "Full Comment Text"
	Channel = ##Heart Rate Channel##
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
## Turn Remaining Channels off ##

	' Begin DataPadColumnSetup
	Column = 6
	FunctionType = "{00000000-0000-0000-0000-000000000000}-0"
	Channel = ##Heart Rate Channel##
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 7
	FunctionType = "{00000000-0000-0000-0000-000000000000}-0"
	Channel = ##Heart Rate Channel##
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 8
	FunctionType = "{00000000-0000-0000-0000-000000000000}-0"
	Channel = ##Heart Rate Channel##
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 9
	FunctionType = "{00000000-0000-0000-0000-000000000000}-0"
	Channel = 8
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 10
	FunctionType = "{00000000-0000-0000-0000-000000000000}-0"
	Channel = 9
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 11
	FunctionType = "{00000000-0000-0000-0000-000000000000}-0"
	Channel = 10
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 12
	FunctionType = "{00000000-0000-0000-0000-000000000000}-0"
	Channel = 11
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup
	
	' Begin DataPadColumnSetup
	Column = 13
	FunctionType = "{00000000-0000-0000-0000-000000000000}-0"
	Channel = 12
	RecordMode = 1
	Options = ""
	Call Doc.DataPadColumnSetup (Column, FunctionType, Channel, RecordMode, Options)
	' End DataPadColumnSetup

	Call Doc.OpenCloseWindow ("Data Pad", 1, False)
	Call Doc.SetViewState ("Data Pad", 1, 61728)
		
## Set cursor to beginning of data ##

	' Begin SetSelection
	Set selobj = CreateObject("ADIChart.Selection")
	Call selobj.SetSelectionRange (0, 0, 0, 1)
	Call selobj.SetChannelRange (0, 1, -1)
	Call selobj.SetChannelRange (1, 1, -1)
	Call selobj.SetChannelRange (2, 1, -1)
	Call selobj.SetChannelRange (3, 1, -1)
	Call selobj.SetChannelRange (4, 1, -1)
	Call selobj.SetChannelRange (5, 1, -1)
	Call selobj.SetChannelRange (6, 1, -1)
	Call selobj.SetChannelRange (7, 1, -1)
	Call selobj.SetChannelRange (8, 1, -1)
	Doc.SelectionObject = selobj
	' End SetSelection
	
	' class SimpleTypeMessage<class TypeIdClass<&struct _GUID const ADICOMObj_GUID,5>,unsigned long>
	Call Doc.PlayMessage ("0x0400000001000000FFFFFFFF01000000FFFFFFFF501B0000AAAA1555010000000500FF7F904BA734BC0DD311B870008048C36FE8000000000100FF7F37F567CC2AD3C44081B60BA26908ABED000000000100000008000000")
	' Begin PositionWindow
	ViewTypeId = "Chart View"
	ViewInstance = 1
	Dim Position(3)
	Position(0) = -105
	Position(1) = 14
	Position(2) = 1939
	Position(3) = 1005
	Call Doc.PositionWindow (ViewTypeId, ViewInstance, Position)
	' End PositionWindow
	
	Call Doc.SetViewState ("Chart View", 1, 61488)

## Find comment or point in data file for data to being ##

	' Begin Find
	ChannelIndex = ##Heart Rate Channel##
	SetAction = kSetActivePoint
	SelectMode = kSelectAround
	SelectTime = 1
	DataDisplayMode = kViewDataVisible
	SelectAll = False
	Direction = kSearchForward
	FindType = "Search for comment"
	FindData = "JustThisChannel=0;WhatToLookFor=##Comment Name##;"
	Call Doc.Find (ChannelIndex, SetAction, SelectMode, SelectTime, DataDisplayMode, SelectAll, Direction, FindType, FindData)
	' End Find
	
## Final local maximum for R spike of QRS interval ##	

	' Begin Find
	ChannelIndex = 7
	SetAction = kSetActivePoint
	SelectMode = kSelectAround
	SelectTime = 1
	DataDisplayMode = kViewDataVisible
	SelectAll = False
	Direction = kSearchForward
	FindType = "Local maxima"
	FindData = "NoiseThreshold=0.05;"
	Call Doc.Find (ChannelIndex, SetAction, SelectMode, SelectTime, DataDisplayMode, SelectAll, Direction, FindType, FindData)
	' End Find

## Set number of repitions for macro ##

	For i = 1 to ##Number for amount of times for macro to repeat##
		
		
		' Begin Find
		ChannelIndex = ##Heart Rate Channel##
		SetAction = kSetToPreviousPoint
		SelectMode = kSelectAround
		SelectTime = 1
		DataDisplayMode = kViewDataVisible
		SelectAll = False
		Direction = kSearchForward
		FindType = "Local maxima"
		FindData = "NoiseThreshold=0.05;"
		Call Doc.Find (ChannelIndex, SetAction, SelectMode, SelectTime, DataDisplayMode, SelectAll, Direction, FindType, FindData)
		' End Find
		
		' The function below will return true if the last operation failed, which will cause the current loop to exit
		Call Doc.AddToDataPad ()
	Next
	Call Doc.OpenView ("Data Pad")
	Call Doc.SetViewState ("Chart View", 1, 61728)


End Sub

