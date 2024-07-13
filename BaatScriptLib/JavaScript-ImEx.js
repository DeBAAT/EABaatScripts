//[group=BaatScriptLib]
!INC EAScriptLib.JavaScript-Logging
!INC EAScriptLib.JavaScript-TaggedValue


/**
 * @file JavaScript-ImEx
 * This script library contains helper functions to assist with IMEX Import and Export of
 * Enterprise Architect Elements and Connectors using an Excel Workbook file.
 * This functionality is similar to the eaexcelimporter of Geert Bellekens. 
 * 
 * Functions in this library are split into three parts: Workbooks, Import and Export.
 * Functions that assist with IMEX Workbooks are prefixed with IMEXW,
 * IMEX Import Functions are prefixed with IMEXI and IMEX Export Functions are prefixed with IMEXE.
 *
 * IMEX Import can be performed by calling the function IMEXIHandleExcelImport().
 * IMEXIHandleExcelImport() requires the name of the WorkSheet containing the information and
 * that the function OnExcelRowImported() is defined in the user's script to be used as a callback
 * whenever row data is read from the IMEX file. The user defined OnExcelRowImported() can query for 
 * information about the current row through the functions EXCELIContainsColumn(), 
 * EXCELIGetColumnValueByName() and EXCELIGetColumnValueByNumber().
 *
 * To perform an IMEX export, the user must call IMEXEHandleExcelExport() which starts an export 
 * session. The call to IMEXEHandleExcelExport() specifies the file name to export to, and the set of 
 * columns that will be exported. Once the session has been initialized with a call to 
 * IMEXEExportInitialize(), the user may continually call IMEXEExportRow() to export a row to file. 
 * Once all rows have been added, the export session is closed by calling IMEXEExportFinalize(). 
 *
 * @author J. de Baat, based on JavaScript - CSV by Sparx Systems
 * @date 2024-07-13
 */

const strGlobalEAPackageName        = "ImEx Package";
const strGlobalEADiagramName        = "ImEx Diagram";
const strDefaultElementsFileName    = "ImExElements.xlsx";
const strDefaultConnectorsFileName  = "ImExConnectors.xlsx";
const strDefaultElementsSheetName   = "ImEx_Elements";
const strDefaultConnectorsSheetName = "ImEx_Connectors";
const strTaggedValuesPrefix         = "TAG_";
const strNoClientConnector          = "ImEx_NoClientConnector";

const strImExStartPackage           = "Package";
const strImExStartDiagram           = "Diagram";
let   strImExStart                  = "None";

const ConnectorArchiMateAssociation = 0;
const ConnectorArchiMateAggregation = 1;
const ConnectorArchiMateComposition = 2;
const ConnectorDirectionUnspecified = "Unspecified";

const strArchiMatePrefix            = "ArchiMate_";
const strArchiMate3Prefix           = "ArchiMate3::";

/**
 * objGlobalEAPackage needs to have a objGlobalEADiagram and visa versa.
 */
var objGlobalEAPackage as EA.Package;
var objGlobalEADiagram as EA.Diagram;

/**
 * If theValue starts with thePrefix then return value without prefix else return empty string
 *
 * @param[in] theValue (string) The value to test.
 * @param[in] thePrefix (string) The prefix to test against.
 */
function IMEXGGetValueWithoutPrefix( theValue, thePrefix )
{

	try {

		const thePrefixLength = thePrefix.length;
		// Session.Output("IMEXGGetValueWithoutPrefix started with thePrefix : " + thePrefix + ", length = " + thePrefixLength + " !" );

		//	If theValue starts with thePrefix then return value without prefix
		if ( theValue.substring( 0, thePrefixLength ).toLowerCase() == thePrefix.toLowerCase() ) {
			// Session.Output( "IMEXGGetValueWithoutPrefix found theValue.substring(" + thePrefixLength + ")= " + theValue.substring( thePrefixLength ) + " !" );
			return theValue.substring( thePrefixLength );
		}
	} catch (err) {
		LOGError( "IMEXGGetValueWithoutPrefix catched error " + err.message + "!" );
		return "";
	}

	return "";

}

/*
 * Check whether theStereotype starts with ArchiMate_ and should be prefixed with ArchiMate3::
 */
function IMEXGCheckArchiMateStereotype( theStereotype )
{

	try {

		let curStereotype = theStereotype;

		// Check whether theStereotype is available
		if ( curStereotype == null ) {
			// Session.Output("IMEXGCheckArchiMateStereotype( " + theStereotype + " ) DID NOT update theStereotype to " + curStereotype + "!!!" );
			return curStereotype;
		}

		// Check whether theStereotype should be prefixed
		if ( curStereotype.substring( 0, strArchiMatePrefix.length ).toLowerCase() == strArchiMatePrefix.toLowerCase() ) {
			curStereotype = strArchiMate3Prefix + curStereotype;
			// Session.Output("IMEXGCheckArchiMateStereotype( " + theStereotype + " ) updated theStereotype to " + curStereotype + "!!!" );
		} else {
			// Session.Output("IMEXGCheckArchiMateStereotype( " + theStereotype + " ) DID NOT update theStereotype to " + curStereotype + "!!!" );
		}

		return curStereotype;

	} catch (err) {
		LOGError( "IMEXGCheckArchiMateStereotype catched error " + err.message + "!" );
		return theStereotype;
	}
}

/*
 * Check theStereotype to define the ClientEnd.Aggregation
 */
function IMEXGCheckArchiMateStereotypeConnector( theStereotype )
{

	// Check whether theStereotype is available
	if ( theStereotype == null ) {
		return ConnectorArchiMateAssociation;
	}

	// Check what to use as Starting point for export
	switch ( theStereotype ) {
		case "ArchiMate_Aggregation" :
		case "ArchiMate3::ArchiMate_Aggregation" :
		{
			// Set to show as ArchiMate_Aggregation
			return ConnectorArchiMateAggregation;
			break;
		}
		case "ArchiMate_Composition" :
		case "ArchiMate3::ArchiMate_Composition" :
		{
			// Set to show as ArchiMate_Composition
			return ConnectorArchiMateComposition;
			break;
		}
		default:
		{
			// Set to show as default Association
			return ConnectorArchiMateAssociation;
			break;
		}
	}

	return ConnectorArchiMateAssociation;
}

/*
 * Check theStereotype to define the ClientEnd.Aggregation
 */
function IMEXGCheckConnectorDirection( theConnector, theDirection )
{

	try {

		// Cast theConnector to EA.Connector so we get intellisense
		var curConnector       as EA.Connector;
		curConnector            = theConnector;

		// Check whether theConnector and theStereotype are available
		if ( ( theConnector == null ) || ( theDirection == null ) ) {
			return;
		}

		// Check what to use as Starting point for export
		switch ( theDirection ) {
			case "Source -> Destination" :
			{
				// Set the SupplierEnd as Navigable
				curConnector.ClientEnd.IsNavigable = false;
				curConnector.ClientEnd.Update();
				curConnector.SupplierEnd.IsNavigable = true;
				curConnector.SupplierEnd.Update();
				break;
			}
			case "Destination -> Source" :
			{
				// Set the ClientEnd as Navigable
				curConnector.ClientEnd.IsNavigable = true;
				curConnector.ClientEnd.Update();
				curConnector.SupplierEnd.IsNavigable = false;
				curConnector.SupplierEnd.Update();
				break;
			}
			case "Bi-Directional" :
			{
				// Set both the ClientEnd and SupplierEnd as Navigable
				curConnector.ClientEnd.IsNavigable = true;
				curConnector.ClientEnd.Update();
				curConnector.SupplierEnd.IsNavigable = true;
				curConnector.SupplierEnd.Update();
				break;
			}
			case "Unspecified" :
			{
				// Set both the ClientEnd and SupplierEnd as Navigable
				curConnector.ClientEnd.IsNavigable = false;
				curConnector.ClientEnd.Update();
				curConnector.SupplierEnd.IsNavigable = false;
				curConnector.SupplierEnd.Update();
				break;
			}
			default:
			{
				// Do nothing
				break;
			}
		}
		// Session.Output( "IMEXGCheckConnectorDirection curConnector(" + curConnector.Name + ").Direction set to " + curConnector.Direction + " based on " + EXCELIGetColumnValueByName("Direction") + ", Navigable: ClientEnd= " + curConnector.ClientEnd.Navigable + ", SupplierEnd= " + curConnector.SupplierEnd.Navigable + "!!!" );

	} catch (err) {
		LOGError( "IMEXGCheckConnectorDirection catched error " + err.message + "!" );
	}
}

/*
 * Add a new PackageDiagram if it is not defined yet
 */
function IMEXGCheckOrAddPackageDiagram( thePackage, thePackageDiagramName )
{

	try {

		// Validate input parameters
		if ( ( thePackage != null ) && ( thePackageDiagramName != "" ) ) {

			var curPackage  as EA.Package;
			var curDiagrams as EA.Collection;
			var curDiagram  as EA.Diagram;

			// Check all Diagrams in thePackage whether the requested thePackageDiagramName already exists
			curPackage  = thePackage;
			curDiagrams = curPackage.Diagrams;
			curDiagram  = curDiagrams.GetByName( thePackageDiagramName );

			// If curDiagram is not found, create a new diagram
			if ( curDiagram == null )
			{
				// Session.Output( "CheckOrAddPackageDiagram addNew because not found: " + thePackageDiagramName );
				curDiagram = curDiagrams.AddNew( thePackageDiagramName, "Logical" );
				curDiagram.Notes = thePackageDiagramName + " created by JavaScript-ImEx library.";
				curDiagram.Update();

				curDiagrams.Refresh();
				curPackage.Update();

			}

			Session.Output( "CheckOrAddPackageDiagram found " + curDiagram.Name + " as part of PackageID=" + thePackage.PackageID + " !" );
			return curDiagram;
		} else {
			LOGError( "CheckOrAddPackageDiagram could NOT add PackageDiagram " + thePackageDiagramName + "!" );
		}

	} catch (err) {
		LOGError( "IMEXGCheckOrAddPackageDiagram catched error " + err.message + "!" );
	}

	return null;

}

/*
 * Get and check the global variables
 */
function IMEXGGetAndCheckGlobalVariables()
{

	//	Check objGlobalEAPackage and objGlobalEADiagram
	if ( ( objGlobalEAPackage == null ) && ( objGlobalEADiagram == null ) ) {
		LOGError( "Either objGlobalEAPackage OR objGlobalEADiagram should be available!" );
		return false;
	}

	//	Check objGlobalEADiagram, get it from objGlobalEAPackage which should not be nothing
	if ( objGlobalEADiagram == null ) {
		objGlobalEADiagram = IMEXGCheckOrAddPackageDiagram( objGlobalEAPackage, strGlobalEADiagramName );
		// Session.Output( "GetAndCheckGlobalVariables found " + objGlobalEADiagram.Name + " as part of objGlobalEAPackage " + objGlobalEAPackage.Name + " !" );
	}

	//	Check objGlobalEAPackage, get it from objGlobalEADiagram which should not be nothing
	if ( objGlobalEAPackage == null ) {
		try {
			objGlobalEAPackage = Repository.GetPackageByID( objGlobalEADiagram.PackageID );
			// Session.Output( "GetAndCheckGlobalVariables found " + objGlobalEAPackage.Name + " as parent of objGlobalEADiagram " + objGlobalEADiagram.Name + " !" );
		} catch (err) {
			LOGError( "IMEXGGetAndCheckGlobalVariables catched error " + err.message + "!" );
			objGlobalEAPackage = null;
		}
	}

	//	Check objGlobalEAPackage and objGlobalEADiagram again
	if ( ( objGlobalEAPackage == null ) || ( objGlobalEADiagram == null ) ) {
		LOGError( "Both objGlobalEAPackage AND objGlobalEADiagram should be available!" );
		return false;
	}

	return true;
}

/*
 * Get and check the global variables for a Project Browser Script
 */
function IMEXGGetAndCheckPackageObject()
{

	// Prepare some global variables
	objGlobalEAPackage = null;
	objGlobalEADiagram = null;

	try {

		// Get the type of element selected in the Project Browser
		let treeSelectedType = Repository.GetTreeSelectedItemType();

		// Handling Code: Uncomment any types you wish this script to support
		// NOTE: You can toggle comments on multiple lines that are currently
		// selected with [CTRL]+[SHIFT]+[C].
		switch ( treeSelectedType )
		{
			case otPackage :
			{
				// Code for when a package is selected
				objGlobalEAPackage = Repository.GetTreeSelectedObject();
				strImExStart = strImExStartPackage;
				Session.Output("IMEXGGetAndCheckPackageObject Found Package : " + objGlobalEAPackage.Name + "!" );

				break;
			}
			case otDiagram :
			{
				// Code for when a diagram is selected
				objGlobalEADiagram = Repository.GetTreeSelectedObject();
				strImExStart = strImExStartDiagram;
				Session.Output("IMEXGGetAndCheckPackageObject Found Diagram : " + objGlobalEADiagram.Name + "!" );

				break;
			}
			default:
			{
				// Error message
				// Session.Output( "This script does not support items of this type." );
				Session.Prompt( "This script does not support items of this type.", promptOK );
				return false;
				break;
			}
		}
	} catch (err) {
		LOGError( "IMEXGGetAndCheckPackageObject catched error " + err.message + "!" );
		return false;
	}

	// Get and check the global variables
	const  validGlobalVariables = IMEXGGetAndCheckGlobalVariables();
	return validGlobalVariables;

}

/*
 * Get and check the global variables for a Diagram Script
 */
function IMEXGGetAndCheckDiagram()
{

	try {

		// Get a reference to the current diagram
		objGlobalEADiagram = Repository.GetCurrentDiagram();

		if ( objGlobalEADiagram != null )
		{

			// Prepare some global variables
			objGlobalEAPackage = null;
			strImExStart       = strImExStartDiagram;
			Session.Output("IMEXGGetAndCheckDiagram Found Diagram : " + objGlobalEADiagram.Name + "!" );

			// Get and check the global variables
			const  validGlobalVariables = IMEXGGetAndCheckGlobalVariables();
			return validGlobalVariables;

		}
		else
		{
			Session.Prompt( "This script requires a diagram to be visible.", promptOK)
		}
	} catch (err) {
		LOGError( "IMEXGGetAndCheckDiagram catched error " + err.message + "!" );
	}

	return false;

}

////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
////																							////
////											IMEX IMPORT										////
////																							////
////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////


/*
 * Handle the Excel Application and the WorkSheet to Import data from
 * Returns an empty string or a string containing an error message
 *
 * @param[in] theExcelSheetName (String) The name of the ExcelSheet to get the information from
 */
function IMEXIHandleExcelImport( theExcelSheetName /* : String */ ) /* : String */
{

	// Get the excel application for Importing
	objExcelApplication = EXCELWStartExcelApplication();
	if ( objExcelApplication == null ) {
		return "IMEXIHandleExcelImport could NOT start Excel.Application!";
	}

	// objExcelApplication STARTED
	// Session.Output("ImExImportElements started Excel.Application !" );

	// Get the EXCEL fileName for this Import session, true for readonly.
	let curExcelFileName = EXCELWGetFileName( strDefaultElementsFileName, true );
	if ( ( curExcelFileName == null ) || ( curExcelFileName == "" ) ) {
		return "IMEXIHandleExcelImport could NOT Get curExcelFileName!";
	}

	// Open curExcelWorkBook for this Import session.
	let curExcelWorkBook = EXCELWOpenWorkbook( curExcelFileName );
	if ( curExcelWorkBook == null ) {
		return "IMEXIHandleExcelImport could NOT Open curExcelWorkBook!";
	}

	// Start the EXCEL Import session with the sheet and columns
	EXCELIImportSheet( theExcelSheetName, true );

	// Check what was used as Starting point for Import
	switch ( strImExStart ) {
		case strImExStartPackage :
		{
			// Do nothing
			break;
		}
		case strImExStartDiagram :
		{
			// Reload objGlobalEADiagram to reflect changes
			try {
				Repository.ReloadDiagram( objGlobalEADiagram.DiagramID );
			} catch (err) {
				LOGError( "IMEXIHandleExcelImport catched error " + err.message + "!" );
			}
			break;
		}
		default:
		{
			// Show Error message
			Session.Output( "This script does not support items of this type!" );
		}
	}

	// Close the excelWorkBooks and Excel Application
	EXCELWCloseWorkbooks( false );

	// Stop the Excel application
	EXCELWStopExcelApplication();

	// Success so return empty string
	return "";

}


/**
 * Returns all column names that are considered standard for ImEx export as an array of Strings.
 */
function IMEXIGetStandardElementColumns() /* : Array */
{
	let standardColumns = [];

	standardColumns.push( "Action" );
	standardColumns.push( "CLASSTYPE" );
	standardColumns.push( "CLASSGUID" );
	standardColumns.push( "ownerField" );
	standardColumns.push( "Pos" );
	standardColumns.push( "Name" );
	standardColumns.push( "Stereotype" );
	standardColumns.push( "ElementID" );
	standardColumns.push( "Notes" );
	standardColumns.push( "Alias" );
	standardColumns.push( "Status" );
	standardColumns.push( "Datatype" );
	standardColumns.push( "Multiplicity" );
	standardColumns.push( "Visibility" );

	return standardColumns;
}

/**
 * Sets the properties on the specified Element if there is a corresponding value for them in the 
 * current row.
 *
 * Element properties that are not set by this function include:
 * 	- Read only properties
 * 	- Collection properties
 * 	- Properties that contain relational information (eg IDs/GUIDs of other elements, connectors 
 *	or packages.
 *	- Modified Date property (this is property is automatically overwritten by the automation 
 *	interface when the element is saved)
 *	- Properties that are themselves a comma separated list
 *
 * @param[in] elementForRow (EA.Element) The element whose properties will be set with the current row's 
 * values
 */
function IMEXISetStandardElementFieldValues( elementForRow /* : EA.Element */ ) /* : void */
{
	if ( excelImportIsImporting )
	{
		try {
			var theElement as EA.Element;
			theElement = elementForRow;

			if ( theElement != null )
			{
				if ( EXCELIContainsColumn("Alias") )
				{
					theElement.Alias = EXCELIGetColumnValueByName("Alias");
				}

				if ( EXCELIContainsColumn("CLASSTYPE") )
				{
					theElement.Type = EXCELIGetColumnValueByName("CLASSTYPE");
				}

				if ( EXCELIContainsColumn("Multiplicity") )
				{
					theElement.Multiplicity = EXCELIGetColumnValueByName("Multiplicity");
				}

				if ( EXCELIContainsColumn("Name") )
				{
					// Session.Output( "IMEXISetStandardElementFieldValues theElement.Name = " + theElement.Name + ", set to:" + EXCELIGetColumnValueByName("Name") + "!!!" );
					theElement.Name = EXCELIGetColumnValueByName("Name");
				}

				if ( EXCELIContainsColumn("Notes") )
				{
					// Session.Output( "IMEXISetStandardElementFieldValues theElement.Name = " + theElement.Name + ", Notes set to:" + EXCELIGetColumnValueByName("Notes") + "!!!" );
					theElement.Notes = EXCELIGetColumnValueByName("Notes");
				}

				if ( EXCELIContainsColumn("Status") )
				{
					theElement.Status = EXCELIGetColumnValueByName("Status");
				}

				if ( EXCELIContainsColumn("Stereotype") )
				{
					theElement.Stereotype = IMEXGCheckArchiMateStereotype( EXCELIGetColumnValueByName("Stereotype") );
				}

				if ( EXCELIContainsColumn("Visibility") )
				{
					theElement.Visibility = EXCELIGetColumnValueByName("Visibility");
				}

				// Commit the updated values
				theElement.Update();
			}
		} catch (err) {
			LOGError( "IMEXISetStandardElementFieldValues catched error " + err.message + "!" );
		}
	}
	else
	{
		LOGWarning( "No import currently running. IMEXISetStandardElementFieldValues() should only be called from within OnExcelRowImported()" );		
	}
}

/**
 * Sets the TaggedValues on the specified Element if there is a corresponding value for them in the 
 * current row.
 *
 * @param[in] elementForRow (EA.Element) The element whose properties will be set with the current row's 
 * values
 */
function IMEXISetElementTaggedValues( elementForRow /* : EA.Element */ ) /* : void */
{
	if ( excelImportIsImporting )
	{
		try {
			var curElement     as EA.Element;
			var curElementTag  as EA.TaggedValue;
			var curElementTags as EA.Collection;

			curElement = elementForRow;

			if ( curElement != null )
			{

				// Process all TaggedValues in excelImportColumnTagsMap
				excelImportColumnTagsMap.forEach(function(value, key) {
						// Session.Output( "IMEXISetElementTaggedValues excelImportColumnTagsMap(" + value + "," + key + ") TESTING for curElement " + curElement.Name + "!!!" );

						if ( EXCELIContainsColumn( key ) )
						{
							// If TaggedValue in import, add it to the curElement
							TVSetElementTaggedValue( curElement, value, EXCELIGetColumnValueByName( key ), true );
							// Session.Output( "IMEXISetElementTaggedValues excelImportColumnTagsMap(" + value + "," + key + ") PROCESSING for curElement " + curElement.Name + "!!!" );
						}
					});

				// Commit the updated values
				curElement.Update();
			}
		} catch (err) {
			LOGError( "IMEXISetElementTaggedValues catched error " + err.message + "!" );
		}
	}
	else
	{
		LOGWarning( "No import currently running. IMEXISetElementTaggedValues() should only be called from within OnExcelRowImported()" );		
	}
}

/**
 * Sets the properties on the specified Connector if there is a corresponding value for them in the current row.
 *
 * @param[in] connectorForRow (EA.Connector) The element whose properties will be set with the current row's values
 */
function IMEXISetStandardConnectorFieldValues( connectorForRow /* : EA.Connector */ ) /* : void */
{
	if ( excelImportIsImporting )
	{
		try {

			// Cast theConnector to EA.Connector so we get intellisense
			var curConnector as EA.Connector;
			curConnector      = connectorForRow;

			// Process all elements in the import values found
			if ( curConnector != null )
			{
				if ( EXCELIContainsColumn("Name") )
				{
					// Session.Output( "IMEXISetStandardConnectorFieldValues curConnector.Name = " + curConnector.Name + ", set to:" + EXCELIGetColumnValueByName("Name") + "!!!" );
					curConnector.Name = EXCELIGetColumnValueByName("Name");
				}

				if ( EXCELIContainsColumn("Connector_Type") )
				{
					curConnector.Type = EXCELIGetColumnValueByName("Connector_Type");
				}

				if ( EXCELIContainsColumn("Direction") )
				{
					IMEXGCheckConnectorDirection( curConnector, EXCELIGetColumnValueByName("Direction") );
				}

				if ( EXCELIContainsColumn("Stereotype") )
				{
					// Set both the curConnector.Stereotype and the curConnector.ClientEnd.Aggregation
					curConnector.Stereotype            = EXCELIGetColumnValueByName( "Stereotype" );
					curConnector.ClientEnd.Aggregation = IMEXGCheckArchiMateStereotypeConnector( curConnector.Stereotype );
				}

				if ( EXCELIContainsColumn("Notes") )
				{
					curConnector.Notes = EXCELIGetColumnValueByName("Notes");
				}

				if ( EXCELIContainsColumn("RouteStyle") )
				{
					curConnector.RouteStyle = EXCELIGetColumnValueByName("RouteStyle");
				}

				// Commit the updated values
				curConnector.Update();
			}
		} catch (err) {
			LOGError( "IMEXISetStandardConnectorFieldValues catched error " + err.message + "!" );
		}
	}
	else
	{
		LOGWarning( "No import currently running. IMEXISetStandardElementFieldValues() should only be called from within OnExcelRowImported()" );		
	}
}

/**
 * Checks and sets the Connector_ID of the Start or End of the Connector.
 *
 * @param[in] connectorForRow (EA.Connector) The element whose properties will be set with the current row's values
 * @param[in] connectorColumn (String) The name of the property to check
 * @param[in] elementForRow (EA.Element) The element to check
 */
function IMEXIGetNewConnectorStartOrEnd( connectorForRow /* : EA.Connector */, connectorColumn /* : String */, elementForRow /* : EA.Element */ ) /* : void */
{
	if ( excelImportIsImporting )
	{
		try {

			// Cast theConnector to EA.Connector so we get intellisense
			var curConnector as EA.Connector;
			var curElement   as EA.Element;
			var newElement   as EA.Element;
			curConnector      = connectorForRow;
			curElement        = elementForRow;
			newElement        = null;

			// Process all elements in the import values found
			if ( ( curConnector != null ) && ( curElement != null ) )
			{
				if ( EXCELIContainsColumn( connectorColumn ) )
				{
					let newConnectorElementID = EXCELIGetColumnValueByName( connectorColumn );
					if ( curElement.ElementID != newConnectorElementID ) {
						newElement = GetElementByID( newConnectorElementID );
						return newElement;
					}
					// Session.Output( "IMEXIGetNewConnectorStartOrEnd could NOT find newConnectorElementID = " + newConnectorElementID + " for curConnector.ConnectorID:" + curConnector.ConnectorID + "!!!" );
				}
			}
		} catch (err) {
			LOGError( "IMEXIGetNewConnectorStartOrEnd catched error " + err.message + "!" );
		}
	}
	else
	{
		LOGWarning( "No import currently running. IMEXISetStandardElementFieldValues() should only be called from within OnExcelRowImported()" );		
	}

	return null;

}

////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
////																							////
////											IMEX EXPORT										////
////																							////
////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////

/*
 * Handle the Excel Application and the WorkSheet to Export data to
 * Returns an empty string or a string containing an error message
 *
 * @param[in] theExcelSheetName (String) The name of the ExcelSheet to get the information from
 */
function IMEXEHandleExcelExport( theExcelSheetName /* : String */ ) /* : String */
{

	// Get the excel application for Exporting
	objExcelApplication = EXCELWStartExcelApplication();
	if ( objExcelApplication == null ) {
		return "IMEXEHandleExcelExport could NOT start Excel.Application!";
	}

	// objExcelApplication STARTED
	// Session.Output("IMEXEHandleExcelExport started Excel.Application!" );

	// Get the EXCEL fileName for this Export session, false for not readonly.
	let curExcelFileName = EXCELWGetFileName( strDefaultElementsFileName, false );
	if ( ( curExcelFileName == null ) || ( curExcelFileName == "" ) ) {
		return "IMEXEHandleExcelExport could NOT Get curExcelFileName!";
	}

	// Open curExcelWorkBook for this Export session.
	let curExcelWorkBook = EXCELWOpenWorkbook( curExcelFileName );
	if ( curExcelWorkBook == null ) {
		return "IMEXEHandleExcelExport could NOT Open curExcelWorkBook!";
	}

	// Initialize the EXCEL Export columns
	let curExportColumns = ImExGetStandardObjectColumns();
	if ( curExportColumns == null ) {
		return "IMEXEHandleExcelExport could NOT Get curExportColumns!";
	}

	// Initialize the EXCEL Export session with the sheet and columns
	curExcelWorkSheet = EXCELEExportInitialize( theExcelSheetName, curExportColumns, true );
	if ( curExcelWorkSheet == null ) {
		return "IMEXEHandleExcelExport could NOT Get curExcelWorkSheet!";
	}

	// Check what to use as Starting point for export
	switch ( strImExStart ) {
		case strImExStartPackage :
		{
			// Export all Connectors of the Elements in the selected Package
			ImExExportPackageObjects();
			break;
		}
		case strImExStartDiagram :
		{
			// Export all Connectors of the Elements in the selected Diagram
			ImExExportDiagramObjects();
			break;
		}
		default:
		{
			// Show Error message
			Session.Output( "This script does not support items of this type!" );
		}
	}

	// Finalizes an EXCEL Export session
	EXCELEExportFinalize();

	// Close the excelWorkBooks and Excel Application
	EXCELWCloseWorkbooks( true );

	// Stop the Excel application
	EXCELWStopExcelApplication();

	// Success so return empty string
	return "";

}

/**
 * Creates and returns an empty Value Map.
 */
function IMEXECreateEmptyValueMap() /* : Map */
{
	let valueMap = new Map();
	return valueMap;
}

/**
 * Returns an array of column names considered standard for EA elements. This array can be used
 * as the columns parameter when calling IMEXEExportInitialize()
 *
 * @return an array of column names 
 */
function IMEXEGetStandardElementColumns() /* : Array */
{
	let standardColumns = [];

	standardColumns.push( "Action" );
	standardColumns.push( "CLASSTYPE" );
	standardColumns.push( "CLASSGUID" );
	standardColumns.push( "ownerField" );
	standardColumns.push( "Pos" );
	standardColumns.push( "Name" );
	standardColumns.push( "Stereotype" );
	standardColumns.push( "ElementID" );
	standardColumns.push( "Notes" );
	standardColumns.push( "Alias" );
	standardColumns.push( "Status" );
	standardColumns.push( "Datatype" );
	standardColumns.push( "Multiplicity" );
	standardColumns.push( "Visibility" );

	return standardColumns;
}

/**
 * Returns an array of column names considered standard for EA elements. This array can be used
 * as the columns parameter when calling IMEXEExportInitialize()
 *
 * @return an array of column names 
 */
function IMEXEGetStandardConnectorColumns() /* : Array */
{
	let standardColumns = [];

	standardColumns.push( "Action" );
	standardColumns.push( "CONNECTORGUID" );
	standardColumns.push( "Connector_ID" );
	standardColumns.push( "Name" );
	standardColumns.push( "Connector_Type" );
	standardColumns.push( "Start_Object_ID" );
	standardColumns.push( "End_Object_ID" );
	standardColumns.push( "Direction" );
	standardColumns.push( "Stereotype" );
	standardColumns.push( "ClientEndAggregation" );
	standardColumns.push( "Notes" );
	standardColumns.push( "RouteStyle" );
	standardColumns.push( "CSO_CLASSTYPE" );
	standardColumns.push( "CSO_CLASSGUID" );
	standardColumns.push( "CSO_OBJECT_ID" );
	standardColumns.push( "CSO_Name" );
	standardColumns.push( "CSO_Stereotype" );
	standardColumns.push( "CSO_Notes" );
	standardColumns.push( "CTO_CLASSTYPE" );
	standardColumns.push( "CTO_CLASSGUID" );
	standardColumns.push( "CTO_OBJECT_ID" );
	standardColumns.push( "CTO_Name" );
	standardColumns.push( "CTO_Stereotype" );
	standardColumns.push( "CTO_Notes" );

	return standardColumns;
}

/**
 * Creates a Value Map of standard property names/values for the specified element. This Value Map 
 * can be used as the valueMap parameter when calling the ExportRow() function.
 *
 * @param[in] element (EA.Element) The element to compile the Value Map for
 *
 * @return A Value Map populated with the provided element's values.
 */
function IMEXEGetStandardElementFieldValues( element /* : EA.Element */ ) /* : Map */
{

	let valueMap = IMEXECreateEmptyValueMap();

	try {

		var theElement as EA.Element;
		theElement = element;

		valueMap.set( "Action", "" );
		valueMap.set( "CLASSTYPE", theElement.Type );
		valueMap.set( "CLASSGUID", theElement.ElementGUID );
		valueMap.set( "ownerField", theElement.ElementGUID );
		valueMap.set( "Pos", -1 );
		valueMap.set( "Name", theElement.Name );
		valueMap.set( "Stereotype", theElement.StereotypeEx );
		valueMap.set( "ElementID", theElement.ElementID );
		valueMap.set( "Notes", theElement.Notes );
		valueMap.set( "Alias", theElement.Alias );
		valueMap.set( "Status", theElement.Status );
		valueMap.set( "Datatype", "" );
		valueMap.set( "Multiplicity", theElement.Multiplicity );
		valueMap.set( "Visibility", theElement.Visibility );

	} catch (err) {
		LOGError( "IMEXEGetStandardElementFieldValues catched error " + err.message + "!" );
		valueMap = null;
	}

	return valueMap;
}

/**
 * Creates a Value Map of standard property names/values for the specified element. This Value Map 
 * can be used as the valueMap parameter when calling the ExportRow() function.
 *
 * @param[in] element (EA.Element) The element to compile the Value Map for
 *
 * @return A Value Map populated with the provided element's values.
 */
function IMEXEGetStandardConnectorFieldValues( element /* : EA.Element */, connector /* : EA.Connector */ ) /* : Map */
{

	let valueMap = IMEXECreateEmptyValueMap();

	try {

		var theConnector       as EA.Connector;
		var theElementClient   as EA.Element;
		var theElementSupplier as EA.Element;
		theElementClient        = element;
		theConnector            = connector;

		// Validate input
		if ( theConnector.ClientID != theElementClient.ElementID ) {
			return strNoClientConnector;
		}
		theElementSupplier = GetElementByID( theConnector.SupplierID );
		if ( theElementSupplier == null ) {
			return null;
		}

		valueMap.set( "Action", "" );
		valueMap.set( "CONNECTORGUID", theConnector.ConnectorGUID );
		valueMap.set( "Connector_ID", theConnector.ConnectorID );
		valueMap.set( "Name", theConnector.Name );
		valueMap.set( "Connector_Type", theConnector.Type );
		valueMap.set( "Start_Object_ID", theConnector.ClientID );
		valueMap.set( "End_Object_ID", theConnector.SupplierID );
		valueMap.set( "Direction", theConnector.Direction );
		valueMap.set( "Stereotype", theConnector.Stereotype );
		valueMap.set( "ClientEndAggregation", theConnector.ClientEnd.Aggregation );
		valueMap.set( "Notes", theConnector.Notes );
		valueMap.set( "RouteStyle", theConnector.RouteStyle );
		valueMap.set( "CSO_CLASSTYPE", theElementClient.Type );
		valueMap.set( "CSO_CLASSGUID", theElementClient.ElementGUID );
		valueMap.set( "CSO_OBJECT_ID", theElementClient.ElementID );
		valueMap.set( "CSO_Name", theElementClient.Name );
		valueMap.set( "CSO_OwnerField", theElementClient.ElementGUID );
		valueMap.set( "CSO_Stereotype", theElementClient.Stereotype );
		valueMap.set( "CSO_Notes", theElementClient.Notes );
		valueMap.set( "CTO_CLASSTYPE", theElementSupplier.Type );
		valueMap.set( "CTO_CLASSGUID", theElementSupplier.ElementGUID );
		valueMap.set( "CTO_OBJECT_ID", theElementSupplier.ElementID );
		valueMap.set( "CTO_Name", theElementSupplier.Name );
		valueMap.set( "CTO_OwnerField", theElementSupplier.ElementGUID );
		valueMap.set( "CTO_Stereotype", theElementSupplier.Stereotype );
		valueMap.set( "CTO_Notes", theElementSupplier.Notes );

	} catch (err) {
		LOGError( "IMEXEGetStandardConnectorFieldValues catched error " + err.message + "!" );
		valueMap = null;
	}

	return valueMap;
}

/**
 * Creates a Value Map of standard property names/values for the specified element. This Value Map 
 * can be used as the valueMap parameter when calling the ExportRow() function.
 *
 * @param[in] element (EA.Element) The element to compile the Value Map for
 *
 * @return A Value Map populated with the provided element's values.
 */
function IMEXEGetElementTaggedValues( map /* : Map */, element /* : EA.Element */ ) /* : Map */
{

	let valueMap = map;

	try {

		var theElement     as EA.Element;
		var curElementTag  as EA.TaggedValue;
		var curElementTags as EA.Collection;

		theElement = element;

		// Process all curElementTags in theElement
		curElementTags = theElement.TaggedValues;
		for ( let i = 0 ; i < curElementTags.Count ; i++ )
		{
			curElementTag = curElementTags.GetAt( i );
			valueMap.set( strTaggedValuesPrefix + curElementTag.Name, curElementTag.Value );
		}

	} catch (err) {
		LOGError( "IMEXEGetElementTaggedValues catched error " + err.message + "!" );
		valueMap = map;
	}

	return valueMap;
}

/**
 * Creates a Value Map of standard property names/values for the specified attribute. This Value Map 
 * can be used as the valueMap parameter when calling the ExportRow() function.
 *
 * @param[in] element (EA.Element) The element to compile the Value Map for
 *
 * @return A Value Map populated with the provided element's values.
 */
function IMEXEGetStandardAttributeFieldValues( element /* : EA.Element */, attribute /* : EA.Attribute */ ) /* : Map */
{

	let valueMap = IMEXECreateEmptyValueMap();

	try {

		var theElement   as EA.Element;
		var theAttribute as EA.Attribute;
		theElement   = element;
		theAttribute = attribute;

		valueMap.set( "Action", "" );
		valueMap.set( "CLASSTYPE", "Attribute" );
		valueMap.set( "CLASSGUID", theAttribute.AttributeGUID );
		valueMap.set( "ownerField", theElement.ElementGUID );
		valueMap.set( "Pos", theAttribute.Pos );
		valueMap.set( "Name", theAttribute.Name );
		valueMap.set( "Stereotype", theAttribute.StereotypeEx );
		valueMap.set( "ElementID", theAttribute.AttributeID );
		valueMap.set( "Notes", theAttribute.Notes );
		valueMap.set( "Alias", theAttribute.Style );
		valueMap.set( "Status", "" );
		valueMap.set( "Datatype", theAttribute.Type );
		valueMap.set( "Multiplicity", theAttribute.LowerBound + ".." + theAttribute.UpperBound );
		valueMap.set( "Visibility", theAttribute.Visibility );

	} catch (err) {
		LOGError( "IMEXEGetStandardAttributeFieldValues catched error " + err.message + "!" );
		valueMap = null;
	}

	return valueMap;
}
