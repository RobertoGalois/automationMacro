REM  *****  BASIC  *****


sub Main

	selectAllNStyle
	correctPlacingPolite

end sub



sub correctPlacingPolite

	Dim oReplace as Object
    oReplace = ThisComponent.createReplaceDescriptor()
    rem oReplace.setSearchString( "( |\t|)+Bien confraternellement." )
    rem oReplace.setReplaceString("\tBien confraternellement.")
    
    oReplace.setSearchString( "^$" )
    oReplace.setReplaceString("")
    
    oReplace.SearchRegularExpression = True
    ThisComponent.replaceAll( oReplace )

end sub



sub selectAllNStyle

	rem ----------------------------------------------------------------------
	rem define variables
	dim document   as object
	dim dispatcher as object
	rem ----------------------------------------------------------------------
	rem get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:SelectAll", "", 0, Array())
	
	rem ----------------------------------------------------------------------
	dim args2(0) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "Bold"
	args2(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args2())
	
	rem ----------------------------------------------------------------------
	dim args3(0) as new com.sun.star.beans.PropertyValue
	args3(0).Name = "Bold"
	args3(0).Value = false
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args3())
	
	rem ----------------------------------------------------------------------
	dim args4(4) as new com.sun.star.beans.PropertyValue
	args4(0).Name = "CharFontName.StyleName"
	args4(0).Value = ""
	args4(1).Name = "CharFontName.Pitch"
	args4(1).Value = 2
	args4(2).Name = "CharFontName.CharSet"
	args4(2).Value = 0
	args4(3).Name = "CharFontName.Family"
	args4(3).Value = 5
	args4(4).Name = "CharFontName.FamilyName"
	args4(4).Value = "Calibri"
	
	dispatcher.executeDispatch(document, ".uno:CharFontName", "", 0, args4())
	
	rem ----------------------------------------------------------------------
	dim args5(2) as new com.sun.star.beans.PropertyValue
	args5(0).Name = "FontHeight.Height"
	args5(0).Value = 11
	args5(1).Name = "FontHeight.Prop"
	args5(1).Value = 100
	args5(2).Name = "FontHeight.Diff"
	args5(2).Value = 0
	
	dispatcher.executeDispatch(document, ".uno:FontHeight", "", 0, args5())

end sub
