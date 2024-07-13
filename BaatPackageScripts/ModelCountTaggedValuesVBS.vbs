'[group=BaatPackageScripts]
option explicit

!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-XML

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name:	ModelCountTaggedValuesVBS
' Author:		J de Baat
' Purpose:		Count all occurrences of all TaggedValues found for PropertyType enums
' Date:			12-12-2023
'
' Note:		Open Excel file for writing and read current contents
' 				Create new Workbook for Parameters
' 				For all PropertyTypes to count
' 					Find all TaggedValues and count them
' 					If  Workbook exists
' 						then Read first row with Headers already found
' 						else Create new Workbook for PropertyType
' 							 Create first row with Headers found
' 					Insert new row with all TaggedValues counted

' Global variables
dim ExcelApp
dim ExcelFileName
dim ExcelWorkBooks
dim i, j

' Global const definitions for use with Excel
const xlCenter 					= -4108
const xlLeft 					= -4131
const xlBelow 					= 1
const xlAbove 					= 0

Class TaggedValueCounts
	'private variables
	Private m_TaggedValue
	Private m_TaggedValueProperty
	Private m_Counts
	Private m_Column

	'constructor
	Private Sub Class_Initialize
		m_TaggedValue = ""
		m_TaggedValueProperty = ""
		m_Counts = ""
		m_Column = 0
	End Sub
	
	'public properties
	
	' TaggedValue property.
	Public Property Get TaggedValue
		TaggedValue = m_TaggedValue
	End Property
	Public Property Let TaggedValue(value)
		m_TaggedValue = value
	End Property
	
	' TaggedValueProperty property.
	Public Property Get TaggedValueProperty
		TaggedValueProperty = m_TaggedValueProperty
	End Property
	Public Property Let TaggedValueProperty(value)
		m_TaggedValueProperty = value
	End Property
	
	' Counts property.
	Public Property Get Counts
		Counts = m_Counts
	End Property
	Public Property Let Counts(value)
		m_Counts = value
	End Property
	
	' Column property.
	Public Property Get Column
		Column = m_Column
	End Property
	Public Property Let Column(value)
		m_Column = value
	End Property

End Class

Dim curTaggedValueCounts

'
' Count all the TaggedValues registered for theTaggedValueProperty
'
Function GetTaggedValuesCountsValues( theTaggedValueProperty )
	dim strSQLQuery
	dim sqlResponse

	strSQLQuery = "select t_objectproperties.Value as TaggedValues, count(*) as Counts from t_objectproperties" & _
                  " where t_objectproperties.Property = '" & theTaggedValueProperty & "'" & _
                  " group by t_objectproperties.Value" & _
                  " order by t_objectproperties.Value; "
	sqlResponse = Repository.SQLQuery( strSQLQuery )
'	Session.Output("strSQLQuery found sqlResponse= " + sqlResponse + "!!!" )

	if len(sqlResponse) > 0 then
		set GetTaggedValuesCountsValues = convertQueryResultToDictionary( sqlResponse, "TaggedValues", "Counts" )
	else
		set GetTaggedValuesCountsValues = nothing
	end if

end function

'
' Get a new object for TaggedValueCounts with the parameter given
'
Function getTaggedValueCounts( theTaggedValue, theTaggedValueProperty, theCounts, theColumn )

    Dim theTaggedValueCounts

	set theTaggedValueCounts = new TaggedValueCounts
	theTaggedValueCounts.TaggedValue = theTaggedValue
    theTaggedValueCounts.TaggedValueProperty = theTaggedValueProperty
    theTaggedValueCounts.Counts = theCounts
    theTaggedValueCounts.Column = theColumn

	set getTaggedValueCounts = theTaggedValueCounts

end function

Function convertQueryResultToDictionary(xmlQueryResult, theKey, theValue)
	dim curColumn
	dim curTotal
    Dim resultDictionary
	set resultDictionary = CreateObject("Scripting.Dictionary")
	resultDictionary.CompareMode = vbBinaryCompare

	curColumn = 1
	curTotal = 0

	set curTaggedValueCounts = getTaggedValueCounts( "ReferentieDatum", "Header", getFormattedDate(), curColumn )
	curColumn = curColumn + 1
	resultDictionary.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts
	set curTaggedValueCounts = getTaggedValueCounts( "ReferentieDagVanWeek", "Header", Weekday(now()), curColumn )
	curColumn = curColumn + 1
	resultDictionary.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts
	set curTaggedValueCounts = getTaggedValueCounts( "ReferentieDagVanMaand", "Header", Day(now()), curColumn )
	curColumn = curColumn + 1
	resultDictionary.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts

	Dim xDoc 
    Set xDoc = CreateObject( "MSXML2.DOMDocument" )
    'load the resultset in the xml document
    If xDoc.LoadXML(xmlQueryResult) Then        
		'select the rows
		Dim rowList
		Set rowList = xDoc.SelectNodes("//Row")
		Dim rowNode
		Dim fieldNode
		'loop rows and find fields
		For Each rowNode In rowList
			dim curKey, curValue
			'loop the field nodes
			For Each fieldNode In rowNode.ChildNodes
				'add the contents
				Select Case fieldNode.nodeName
					Case theKey
						curKey = fieldNode.Text
					Case theValue
						curValue = fieldNode.Text
						curTotal = curTotal + curValue
				End Select
			Next
			' Add the found key,value pair to the resultDictionary
			if len(curKey) > 0 then
				set curTaggedValueCounts = getTaggedValueCounts( curKey, "", curValue, curColumn )
				curColumn = curColumn + 1
				resultDictionary.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts
				' Session.Output("convertQueryResultToDictionary found resultDictionary( " & curTaggedValueCounts.TaggedValue & " , " & curTaggedValueCounts.Counts & " )!" )
			end if
		Next

		' Add the found curTotal to the resultDictionary
		if curTotal > 0 then
			set curTaggedValueCounts = getTaggedValueCounts( "Totaal", "", curTotal, curColumn )
			curColumn = curColumn + 1
			resultDictionary.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts
			' Session.Output("convertQueryResultToDictionary found resultDictionary( " & curTaggedValueCounts.TaggedValue & " , " & curTaggedValueCounts.Counts & " )!" )
		end if
	end if

	' Only return the values found when there are values found
	if curTotal > 0 then
		set convertQueryResultToDictionary = resultDictionary
	else
		set convertQueryResultToDictionary = nothing
	end if

end function


'
' Get the current date in the format yyyy-mm-dd
'
Function getFormattedDate()
	getFormattedDate = Year(now()) & "-" & Right("0" & Month(now()),2) & "-" & Right("0" & Day(now()),2)
end function

'
' Get the current time in the format hh:mm:ss
'
Function getFormattedTime( separator )
	getFormattedTime = Right("0" & Hour(now()),2) & separator & Right("0" & Minute(now()),2) & separator & Right("0" & Second(now()),2)
end function

'
' Process the current propertyTypeWorksheet
'
Function processPropertyTypeWorksheet( thePropertyType, thePropertyTypeIndex )

	dim curPropertyType As EA.PropertyType
	set curPropertyType = thePropertyType

	' Get the propertyTypeWorksheet for the requested curPropertyType.Tag (max 30 characters)
	dim propertyTypeWorksheet
	dim propertyTypeHeaders
	set propertyTypeWorksheet = getWorksheet( Left( curPropertyType.Tag, 30 ), thePropertyTypeIndex )
	set propertyTypeHeaders = getPropertyTypeHeaders( curPropertyType, propertyTypeWorksheet )

	' Count all the TaggedValues registered for curPropertyType.Tag
	dim propertyTypeCountsValues
	set propertyTypeCountsValues = GetTaggedValuesCountsValues( curPropertyType.Tag )
	processPropertyTypeValues curPropertyType, propertyTypeWorksheet, propertyTypeCountsValues, propertyTypeHeaders 
	if propertyTypeCountsValues is nothing then
		' Session.Output("GetTaggedValuesCounts found NO propertyTypeCountsValues for curPropertyType.Tag( " & curPropertyType.Tag & " )!" )
	else
		' Session.Output("GetTaggedValuesCounts processed " & propertyTypeCountsValues.Count & " propertyTypeCountsValues for curPropertyType.Tag( " & curPropertyType.Tag & " )!" )
		' Session.Output("getPropertyTypeHeaders found " & propertyTypeHeaders.Count & " propertyTypeHeaders for curPropertyType.Tag( " & curPropertyType.Tag & " )!" )
	end if

	' Autofit and center columns in the propertyTypeWorksheet
	dim targetRange
	set targetRange = propertyTypeWorksheet.Range(propertyTypeWorksheet.Cells(1,1), propertyTypeWorksheet.Cells(200,80))
	targetRange.Columns.Autofit
	set targetRange = propertyTypeWorksheet.Range("A:Z")
	targetRange.HorizontalAlignment = xlCenter

end function

'
' Get the PropertyTypeHeaders
'
Function getPropertyTypeHeaders( thePropertyType, thePropertyTypeWorksheet )

	dim curColumn
	dim curPropertyType As EA.PropertyType
	Dim curPropertyTypeEnums
	dim dictHeaders
	curColumn = 1
	set curPropertyType = thePropertyType
	set dictHeaders = CreateObject( "Scripting.Dictionary" )
	dictHeaders.CompareMode = vbBinaryCompare

	' Check whether the first row contains Headers
	if ( thePropertyTypeWorksheet.Cells(1,1).Value = "ReferentieDatum" ) then
		' Get the Headers from the first row
		dim curCol
		dim curColValue
		curColumn = 1
		' Session.Output("getPropertyTypeHeaders found ReferentieDatum in first row for curPropertyType.Tag( " & curPropertyType.Tag & " )!" )
		for each curCol in thePropertyTypeWorksheet.Columns
			curColValue = thePropertyTypeWorksheet.Cells(1,curColumn).Value
			if not dictHeaders.Exists( curColValue ) then
				set curTaggedValueCounts = getTaggedValueCounts( curColValue, curPropertyType.Tag, "0", curColumn )
				dictHeaders.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts
			end if
			curColumn = curColumn + 1
			if not len(thePropertyTypeWorksheet.Cells(1,curColumn).Value) > 0 then
				exit for
			end if
		next
	else
		' Fill the first row columns with Referentie Data
		curColumn = 1
		' Session.Output("getPropertyTypeHeaders NOT FOUND ReferentieDatum in first row for curPropertyType.Tag( " & curPropertyType.Tag & " )!" )
		' set curTaggedValueCounts = getTaggedValueCounts( "ReferentieDatum", curPropertyType.Tag, getFormattedDate() & " - " & getFormattedTime(":"), curColumn )
		set curTaggedValueCounts = getTaggedValueCounts( "ReferentieDatum", curPropertyType.Tag, getFormattedDate(), curColumn )
		dictHeaders.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts
		thePropertyTypeWorksheet.Cells( 1, curColumn ).Value = curTaggedValueCounts.TaggedValue
		curColumn = curColumn + 1
		set curTaggedValueCounts = getTaggedValueCounts( "ReferentieDagVanWeek", curPropertyType.Tag, Weekday(now()), curColumn )
		dictHeaders.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts
		thePropertyTypeWorksheet.Cells( 1, curColumn ).Value = curTaggedValueCounts.TaggedValue
		curColumn = curColumn + 1
		set curTaggedValueCounts = getTaggedValueCounts( "ReferentieDagVanMaand", curPropertyType.Tag, Day(now()), curColumn )
		dictHeaders.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts
		thePropertyTypeWorksheet.Cells( 1, curColumn ).Value = curTaggedValueCounts.TaggedValue
		curColumn = curColumn + 1

		curPropertyTypeEnums = GetPropertyTypeEnums( curPropertyType )
		if not curPropertyTypeEnums(0) = false then
			' Process all curPropertyTypeEnums found
			for i = 0 to Ubound(curPropertyTypeEnums)
				set curTaggedValueCounts = getTaggedValueCounts( curPropertyTypeEnums(i), curPropertyType.Tag, "0", curColumn )
				dictHeaders.Add curTaggedValueCounts.TaggedValue, curTaggedValueCounts
				thePropertyTypeWorksheet.Cells( 1, curColumn ).Value = curTaggedValueCounts.TaggedValue
				curColumn = curColumn + 1
				' Session.Output("getPropertyTypeHeaders for curPropertyType(" & curPropertyType.Tag & ") Processed curPropertyTypeEnums(" & i & ")=" & curPropertyTypeEnums(i) & "!" )
			next
		end if

	end if

	set getPropertyTypeHeaders = dictHeaders

end function

'
' Get the list of Enum values defined in PropertyType.Detail with format:
' 		Type=Enum;
' 		Values=<Value1>,<Value2>,...,<ValueN>;
' 		Default=<Default>;
'
Private Function GetPropertyTypeEnums( thePropertyType )
	Dim startPos, endPos
	Dim curPropertyTypeEnums
	Dim curPropertyTypeValues
	dim curPropertyType As EA.PropertyType

	set curPropertyType = thePropertyType
	curPropertyTypeValues = curPropertyType.Detail

	' Get m_startpos as first position after "Values="
	startPos = InStr(1, curPropertyTypeValues, "Values=", 1) + Len("Values=")
	If startPos > 0 Then
		' Get the part of curPropertyTypeValues after "Values="
		endPos = Len(curPropertyTypeValues)
		curPropertyTypeValues = Mid(curPropertyTypeValues, startPos, endPos - startPos)
		' Strip the part of remaining curPropertyTypeValues after ";"
		endPos = InStr(1, curPropertyTypeValues, ";", 1)
		' Get the Values part of curPropertyTypeValues between "Values=" and ";"
		curPropertyTypeValues = Left(curPropertyTypeValues, endPos - 1)
		' Session.Output("GetPropertyTypeEnums found (" & curPropertyTypeValues & ") in .Detail=" & curPropertyType.Detail & "!" )

		' Split the string in an array of enum values
		curPropertyTypeEnums = Split(curPropertyTypeValues, ",", -1, 1)
		GetPropertyTypeEnums = curPropertyTypeEnums
	Else
		GetPropertyTypeEnums(0) = false
		Session.Output("GetPropertyTypeEnums found NO ENUM VALUES in .Detail=" & curPropertyType.Detail & "!" )
	End If
End Function

'
' Process theValues found and use the PropertyTypeHeaders
'
Function processPropertyTypeValues( thePropertyType, thePropertyTypeWorksheet, theValues, theHeaders )

	dim newColumn
	dim curPropertyType As EA.PropertyType
	dim dictHeaders
	dim allItems
	dim curItem
	dim dictColumn

	set curPropertyType = thePropertyType
	set dictHeaders = theHeaders
	dictHeaders.CompareMode = vbBinaryCompare

	' Insert a new second row to fill theValues found
	' thePropertyTypeWorksheet.Cells(2,2).Value = "TestValue"
	dim targetRange
	set targetRange = thePropertyTypeWorksheet.Range("A2")
	targetRange.EntireRow.Insert

	' Fill the already available header columns with a default value
	allItems = dictHeaders.Items
	for each curItem in allItems
		dictColumn = dictHeaders.Item(curItem.TaggedValue).Column
		thePropertyTypeWorksheet.Cells(2,dictColumn).Value = curItem.Counts
	next

	' Process theValues found before and insert them in the new second row
	if not theValues is nothing then
		allItems = theValues.Items
		newColumn = dictHeaders.Count + 1
		for each curItem in allItems
			' Get an existing Item, even when the Key IsNumeric
			dictColumn = 0
			if dictHeaders.Exists(curItem.TaggedValue) then
				dictColumn = dictHeaders.Item(curItem.TaggedValue).Column
			else
				Session.Output("processPropertyTypeValues testing IsNumeric theValues.Items(" & dictColumn & ")= " & curItem.TaggedValue & "!" )
				if ( IsNumeric( curItem.TaggedValue ) ) then
					if ( dictHeaders.Exists(CInt(curItem.TaggedValue)) ) then
						dictColumn = dictHeaders.Item(CInt(curItem.TaggedValue)).Column
					end if
				end if
			end if
			if ( dictColumn > 0 ) then
				thePropertyTypeWorksheet.Cells(2, dictColumn).Value = curItem.Counts
				if len( thePropertyTypeWorksheet.Cells(1, dictColumn).Value ) = 0 then
					thePropertyTypeWorksheet.Cells(1, dictColumn).Value = curItem.TaggedValue
				end if
				' Session.Output("processPropertyTypeValues only adding theValues.Items(" & dictColumn & ")= " & curItem.TaggedValue & ", Cell=" & thePropertyTypeWorksheet.Cells(1, dictColumn).Value & "!" )
			else
				' Session.Output("processPropertyTypeValues adding header and Cells(1," & newColumn & ")= " & curItem.TaggedValue & "!" )
				dictHeaders.Add curItem.TaggedValue, curItem
				thePropertyTypeWorksheet.Cells(1, newColumn).Value = curItem.TaggedValue
				thePropertyTypeWorksheet.Cells(2, newColumn).Value = curItem.Counts
				newColumn = newColumn + 1
			end if
		next
	end if

end function

'
' Process all PropertyTypes in the Repository
'
Function processAllPropertyTypes()

	' Process all PropertyTypes in the Repository
	dim curPropertyType As EA.PropertyType
	dim curPropertyTypeWorksheet
	curPropertyTypeWorksheet = 2 ' Parameters is Worksheet 1
	for each curPropertyType in Repository.PropertyTypes
		if Left( curPropertyType.Detail, 9 ) = "Type=Enum" then
			processPropertyTypeWorksheet curPropertyType, curPropertyTypeWorksheet
			Session.Output("==> ExcelWorkBooks processed propertyTypeWorksheet[" & curPropertyTypeWorksheet & "]= " & curPropertyType.Tag & " with timestampDate: " & getFormattedDate() & "-" & getFormattedTime(":") & "!" )
			curPropertyTypeWorksheet = curPropertyTypeWorksheet + 1
		else
			Session.Output("===> SKIP processing NO ENUM propertyType[" & curPropertyType.Tag & "] with timestampDate: " & getFormattedDate() & "-" & getFormattedTime(":") & "!" )
		end if
	next

end function

'
' Fill the first sheet with Parameters
'
Function processParameterWorksheet()

	' Get the ParameterSheet
	dim parameterWorksheet
	set parameterWorksheet = getWorksheet( "Parameters", 1 )

	' Create the parameterValues in the parameterWorksheet
	i = 1
	j = 1

	parameterWorksheet.Cells(i,1).Value = "Generatie datum en tijd"
	parameterWorksheet.Cells(i,2).Value = Date() & " -- " & Time()
	i = i + 1
	parameterWorksheet.Cells(i,1).Value = "ReferentieDatum"
	parameterWorksheet.Cells(i,2).Value = getFormattedDate()
	i = i + 1
	parameterWorksheet.Cells(i,1).Value = "ReferentieDagVanWeek"
	parameterWorksheet.Cells(i,2).Value = Weekday(now())
	i = i + 1
	parameterWorksheet.Cells(i,1).Value = "ReferentieDagVanMaand"
	parameterWorksheet.Cells(i,2).Value = Day(now())
	i = i + 1
	Session.Output("ExcelWorkBooks filled parameterWorksheet with timestampDate: " & getFormattedDate() & " - " & getFormattedTime(":") & "!" )

	' Create the list of PropertyTypes in the parameterWorksheet
	dim curPropertyType As EA.PropertyType
	i = i + 1
	parameterWorksheet.Cells(i,1).Value = "PropertyType Overzicht:"
	parameterWorksheet.Cells(i,2).Value = Repository.PropertyTypes.Count
	i = i + 1
	j = 1
	for each curPropertyType in Repository.PropertyTypes
		parameterWorksheet.Cells(i,1).Value = "PropertyType[" & j & "]:"
		parameterWorksheet.Cells(i,2).Value = curPropertyType.Tag
		parameterWorksheet.Cells(i,3).Value = j
		parameterWorksheet.Cells(i,4).Value = curPropertyType.Description
		parameterWorksheet.Cells(i,5).Value = curPropertyType.Detail
		i = i + 1
		j = j + 1
	next

	' Create the list of sheets in the parameterWorksheet
	dim curWorksheet
	i = i + 1
	parameterWorksheet.Cells(i,1).Value = "Worksheet Overzicht:"
	parameterWorksheet.Cells(i,2).Value = ExcelWorkBooks.Sheets.Count
	i = i + 1
	j = 1
	for each curWorksheet in ExcelWorkBooks.Sheets
		parameterWorksheet.Cells(i,1).Value = "Sheet[" & j & "]:"
		parameterWorksheet.Cells(i,2).Value = curWorksheet.Name
		parameterWorksheet.Cells(i,3).Value = j
		i = i + 1
		j = j + 1
	next

	' Autofit columns in the parameterWorksheet
	dim targetRange
	set targetRange = parameterWorksheet.Range(parameterWorksheet.Cells(1,1), parameterWorksheet.Cells(200,80))
	targetRange.Columns.Autofit
	set targetRange = parameterWorksheet.Range("B:C")
	targetRange.HorizontalAlignment = xlCenter

end function

'
' Get the worksheet by name
'
Function getWorksheet( getWorksheetName, beforeSheetIndex )

	' Get the ParameterSheet
	dim curWorksheet
	dim foundWorksheet
	set foundWorksheet = nothing

	' Check if getWorksheetName exists
	for each curWorksheet in ExcelWorkBooks.Sheets
		' Session.Output("ExcelWorkBooks checking curWorksheet.Name: " & curWorksheet.Name & "!" )
		' Check the names of the worksheets in LowerCase and with max 30 chars
		if lcase(curWorksheet.Name) = lcase(Left( getWorksheetName, 30 )) then
			set foundWorksheet = curWorksheet
			foundWorksheet.Name = Left( getWorksheetName, 30 )
			' Session.Output("ExcelWorkBooks found curWorksheet.Name: " & curWorksheet.Name & "!" )
			exit for
		end if
	next

	' If parameterWorksheet not exists yet then create
	if foundWorksheet is nothing then
		' Session.Output("getWorksheet adding [" & getWorksheetName & "] with beforeSheetIndex= " & beforeSheetIndex & ", and ExcelWorkBooks.Sheets.Count= " & ExcelWorkBooks.Sheets.Count & "!" )
		' Check the beforeIndex. In -1 then add after the last one
		if beforeSheetIndex > 0 and beforeSheetIndex <= ExcelWorkBooks.Sheets.Count then
			Set foundWorksheet = ExcelWorkBooks.Sheets.Add(ExcelWorkBooks.Sheets(beforeSheetIndex))
		else
			Set foundWorksheet = ExcelWorkBooks.Sheets.Add(,ExcelWorkBooks.Sheets(ExcelWorkBooks.Sheets.Count))
		end if
		' Set the name of this sheet with max 30 characters
		foundWorksheet.Name = Left( getWorksheetName, 30 )
	else
		' Session.Output("ExcelWorkBooks found foundWorksheet.Name: " & foundWorksheet.Name & "!" )
	end if

	set getWorksheet = foundWorksheet

end function

'
' Get the getExcelWorkbooks file to read from and write the output to
'
Function getExcelWorkbooks()

	dim strDefaultFileName
	' strDefaultFileName = "ModelCountTaggedValuesVBS-" & getFormattedDate() & " " & getFormattedTime( "-" ) & ".xlsx"
	strDefaultFileName = "Actium Repository TaggedValues EA SaaS.xlsx"

	' Get the outputFileName to store the information in
	dim project
	set project = Repository.GetProjectInterface()
	ExcelFileName = project.GetFileNameDialog (strDefaultFileName, "Excel Files|*.xls;*.xlsx;*.xlsm", 1, 2 ,"", 1) 'save as with overwrite prompt: OFN_OVERWRITEPROMPT
	Session.Output( "ExcelFileName = " & ExcelFileName )

	' Get the excelFileWorkBook to store the information in
	dim fileSystemObject
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	if fileSystemObject.FileExists(ExcelFileName) then
		' Open the existing excelFileWorkBook to store the information in
		set ExcelWorkBooks = ExcelApp.Workbooks.Open( ExcelFileName )
	else
		' Make sure we have a filename
		if len(ExcelFileName) = 0 then
			set ExcelWorkBooks = nothing
			Session.Output( "ExcelWorkBooks not created because ExcelFileName is not defined" )
		else
			' Create a new empty excelFileWorkBook to store the information in
			set ExcelWorkBooks = ExcelApp.Workbooks.Add()
			ExcelWorkBooks.ActiveSheet.Name = "Parameters"
			Session.Output("ExcelWorkBooks found ActiveSheet.Name: " & ExcelWorkBooks.ActiveSheet.Name & "!" )
			ExcelWorkBooks.Saveas( ExcelFileName )
		end if
	end if

end function


'
' Project Browser Script main function
'
sub ModelCountTaggedValuesVBS()

	dim numPropertyTypes

	' Show the script output window
	Repository.EnsureOutputVisible "Script"
	set ExcelApp = CreateObject("Excel.Application")
	set curTaggedValueCounts = new TaggedValueCounts

	Session.Output( "ModelCountTaggedValuesVBS" )
	Session.Output( "=======================================" )

	Repository.PropertyTypes.Refresh
	numPropertyTypes = Repository.PropertyTypes.Count

	' Process the numPropertyTypes found
	if ( numPropertyTypes > 0 ) then

		' Get ExcelWorkBooks to read from and write to
		dim excelWorkSheets
		getExcelWorkbooks()

		if ( not ExcelWorkBooks is nothing ) and (len(ExcelApp.Workbooks.Count) > 0 ) then

			' Get excelWorkSheets to process the information
			set excelWorkSheets = ExcelWorkBooks.Sheets
			Session.Output("ExcelApp.Workbooks Found " & ExcelApp.Workbooks.Count & " ExcelWorkBooks and " & excelWorkSheets.Count & " excelWorkSheets!" )
			' Show the worksheets found
			' dim curWorksheet
			' for each curWorksheet in ExcelWorkBooks.Sheets
			' 	Session.Output("ExcelWorkBooks found curWorksheet.Name " & curWorksheet.Name & " in excelWorkSheets!" )
			' next

			' Fill the first sheet with Parameters
			processParameterWorksheet()

			' Process all PropertyTypes in the Repository
			processAllPropertyTypes()

			Session.Prompt "Processed " & numPropertyTypes & " PropertyTypes.", promptOK
		else
			' No valid ExcelApp.Workbooks Found
			Session.Output("Cancelled Processing " & numPropertyTypes & " PropertyTypes because no valid ExcelApp.Workbooks Found in ExcelWorkBooks!" )
			Session.Prompt "Cancelled Processing " & numPropertyTypes & " PropertyTypes.", promptOK
		end if

	else
		Session.Output( "Found NO valid PropertyTypes!!!" )
		Session.Prompt "Found NO valid PropertyTypes!!!", promptOK
	end if

	' Save the active ExcelWorkBooks and Close the ExcelApp
	if ( not ExcelWorkBooks is nothing ) then
		ExcelWorkBooks.Save
	end if
	ExcelApp.Workbooks.Close

	Session.Output( "Processed " & numPropertyTypes & " PropertyTypes and closed " & ExcelFileName & "." )
	Session.Output( "======================================= Closed" )
	
end sub

ModelCountTaggedValuesVBS
