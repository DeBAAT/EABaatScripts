//[group=BaatScriptLib]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx

/*
 * Script Name:	ImExExportConnectors
 * Author:		J de Baat
 * Purpose:		Export the information from Connectors in the selected Package or Diagram
 * Date:		20-12-2024
 * 
 * Note:	    Open Excel file for writing contents
 */

/*
 * Handle the Excel Application and the WorkSheet to export data to
 */
function ImExExportConnectors( )
{

	// Initialize the EXCEL Export session with the sheet and columns
	let curHandleExcelExportResult = IMEXEHandleExcelExport( strDefaultConnectorsSheetName );
	return curHandleExcelExportResult;

}

/*
 * Export the Connectors of the Elements found in this Package
 */
function ImExExportPackageObjects()
{
	let curNumConnectors = 0;

	curNumConnectors = ProcessPackage( objGlobalEAPackage, 0 );
	Session.Output("ImExExportPackageConnectors Processed objGlobalEAPackage( " + objGlobalEAPackage.Name + " ) and newNumConnectors = " + curNumConnectors + "!!!" );

}

/*
 * Export the Connectors of the Elements found in this Diagram
 */
function ImExExportDiagramObjects()
{
	let curNumConnectors = 0;

	curNumConnectors = ProcessDiagram( objGlobalEADiagram, 0 );
	Session.Output("ImExExportDiagramConnectors Processed objGlobalEADiagram( " + objGlobalEADiagram.Name + " ) and newNumConnectors = " + curNumConnectors + "!!!" );

}

/*
 * Initialize the EXCEL Export columns for Connectors
 */
function ImExGetStandardObjectColumns()
{

	// Initialize the EXCEL Export columns for Connectors
	let curExportColumns = IMEXEGetStandardConnectorColumns();
	return curExportColumns;

}

/*
 * Process theElement provided as parameter and its Connectors
 */
function ProcessElementConnectors( theElement, theNumConnectors )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement           as EA.Element;
	var curElementConnector  as EA.Connector;
	var curElementConnectors as EA.Collection;
	var curConnectorTag      as EA.TaggedValue;
	var curConnectorTags     as EA.Collection;
	let curNumConnectors;
	let newNumConnectors;
	let curTagColumnArray = [];

	curElement       = theElement;
	curNumConnectors = theNumConnectors;

	if ( curElement == undefined ) {
		return theNumConnectors;
	}

	// Process all curElementConnectors in curElement
	curElementConnectors = curElement.Connectors;
	let curElementConnectorsCount = curElementConnectors.Count;
	for ( let i = 0 ; i < curElementConnectorsCount ; i++ )
	{
		curElementConnector = curElementConnectors.GetAt( i );
		// newNumConnectors    = ProcessConnector( curElementConnector, curNumConnectors );
		// curNumConnectors    = newNumConnectors;
		// Session.Output("ProcessElementConnectors Processed curElementConnectors(" + i + ") of [" + curElementConnectors.Count + "] and curNumConnectors = " + curNumConnectors + "!!!" );

		// Process all curConnectorTags in curElementConnector
		curConnectorTags = curElementConnector.TaggedValues;
		let curConnectorTagsCount = curConnectorTags.Count;
		for ( let j = 0 ; j < curConnectorTagsCount ; j++ )
		{
			curConnectorTag  = curConnectorTags.GetAt( j );
			curTagColumnArray.push( strTaggedValuesPrefix + curConnectorTag.Name );
			// Session.Output("ProcessElementConnectors Processed curConnectorTags(" + j + ") of [" + curConnectorTags.Count + "] and curNumElements = " + curNumElements + "!!!" );
		}

		// Add the TaggedValues to Export columns
		if ( curTagColumnArray.length > 0 ) {
			EXCELEAddExportColumns( curTagColumnArray );
		}

		// Build the curConnectorMap with values to Export
		let curConnectorMap   = IMEXEGetStandardConnectorFieldValues( curElement, curElementConnector );

		// Export the Connector of this Element
		if ( curConnectorMap == strNoClientConnector ) {
			// Session.Output( "ProcessElementConnectors curConnectorMap found NoClientConnector (i.e. duplicate) for ElementID " + curElement.ElementID + " and Connector.ClientID " + curElementConnector.ClientID + "!" );
			// BLOGWarning( "ProcessElementConnectors curConnectorMap found NoClientConnector (i.e. duplicate) for ElementID " + curElement.ElementID + " and Connector.ClientID " + curElementConnector.ClientID + "!" );
		} else {

			// Export the Connector of this Element
			if ( curConnectorMap != null ) {
				let curTaggedValueMap = IMEXEGetConnectorTaggedValues( curConnectorMap, curElementConnector );
				EXCELEExportRow( curTaggedValueMap );
				curNumConnectors++;
				// Session.Output( "ProcessElementConnectors Processed curTaggedValueMap for ElementID " + curElement.ElementID + " and Connector.ClientID " + curElementConnector.ClientID + "!" );
			} else {
				BLOGError( "ProcessElementConnectors could NOT get curTaggedValueMap for ElementID " + curElement.ElementID + " and Connector.ClientID " + curElementConnector.ClientID + "!" );
			}
		}

		// Clean up memory
		curConnectorMap  = null;
	}

	// Clean up memory
	curElementConnector  = null;
	curElementConnectors = null;
	curConnectorTag      = null;
	curConnectorTags     = null;
	curTaggedValueMap    = null;

	// Session.Output("ProcessElementConnectors Processed Element(" + curElement.Name + ") with ObjectType(Type)=" + curElement.Type + ", Connectors.Count=[" + curElement.Connectors.Count + "] and curNumConnectors = " + curNumConnectors + "!!!" );

	return curNumConnectors;
}

/*
 * Process thePackage provided as parameter and its Elements and SubPackages
 */
function ProcessPackage( thePackage, theNumConnectors )
{

	// Cast thePackage to EA.Package so we get intellisense
	var curPackage         as EA.Package;
	var curPackageElements as EA.Collection;
	var curPackagePackages as EA.Collection;
	var curPackageElement  as EA.Element;
	var curPackagePackage  as EA.Package;
	let curNumConnectors;
	let newNumConnectors;

	curPackage       = thePackage;
	curNumConnectors = theNumConnectors;

	// Process all Elements in curPackage
	curPackageElements = curPackage.Elements;
	for ( let i = 0 ; i < curPackageElements.Count ; i++ )
	{
		curPackageElement = curPackageElements.GetAt( i );
		newNumConnectors  = ProcessElementConnectors( curPackageElement, curNumConnectors );
		curNumConnectors  = newNumConnectors;
		// Session.Output("ProcessPackage Processed curPackageElement(" + i + ") of [" + curPackageElements.Count + "] and curNumConnectors = " + curNumConnectors + "!!!" );
	}

	// Clean up memory
	curPackageElement  = null;
	curPackageElements = null;


	// Process all subPackages in curPackage
	curPackagePackages = curPackage.Packages;
	// Session.Output("ProcessPackage Starting recursively for Package(" + curPackage.Name + ") with curPackagePackages.Count=[" + curPackagePackages.Count + "] and curNumConnectors = " + curNumConnectors + "!!!" );
	for ( let i = 0 ; i < curPackagePackages.Count ; i++ )
	{
		curPackagePackage = curPackagePackages.GetAt( i );
		newNumConnectors  = ProcessPackage( curPackagePackage, curNumConnectors );
		curNumConnectors  = newNumConnectors;
		// Session.Output("curPackagePackage(" + curPackagePackage.Name + ") with Elements.Count=[" + curPackagePackage.Elements.Count + "] and Packages.Count=[" + curPackagePackage.Packages.Count + "] !!!" );
	}

	// Clean up memory
	curPackagePackage  = null;
	curPackagePackages = null;

	Session.Output("ProcessPackage Processed Package(" + curPackage.Name + ") with Elements.Count=[" + curPackage.Elements.Count + "], Packages.Count=[" + curPackage.Packages.Count + "] and curNumConnectors = " + curNumConnectors + "!!!" );

	return curNumConnectors;

}

/*
 * Process theDiagram provided as parameter and its Elements
 */
function ProcessDiagram( theDiagram, theNumConnectors )
{

	// Cast theDiagram to EA.Diagram so we get intellisense
	var curDiagram        as EA.Diagram;
	var curDiagramObjects as EA.Collection;
	var curDiagramObject  as EA.DiagramObject;
	var curDiagramElement as EA.Element;
	let curNumConnectors;
	let newNumConnectors;

	curDiagram     = theDiagram;
	curNumConnectors = theNumConnectors;

	// Process all Elements in curDiagram
	curDiagramObjects = curDiagram.DiagramObjects;
	for ( let i = 0 ; i < curDiagramObjects.Count ; i++ )
	{
		// Get the curDiagramDiagramObject from the Collection
		curDiagramObject = curDiagramObjects.GetAt( i );

		// Get the curDiagramElement using the curDiagramObject.ElementID
		curDiagramElement = Repository.GetElementByID( curDiagramObject.ElementID );
		if ( curDiagramElement != null ) {
			newNumConnectors = ProcessElementConnectors( curDiagramElement, curNumConnectors );
			curNumConnectors = newNumConnectors;
			// Session.Output("ProcessDiagram Processed curDiagramElement(" + i + ") of [" + curDiagramObjects.Count + "] and curNumConnectors = " + curNumConnectors + "!!!" );
		}
	}

	// Clean up memory
	curDiagramElement  = null;
	curDiagramElements = null;

	Session.Output("ProcessDiagram Processed Diagram(" + curDiagram.Name + ") with DiagramObjects.Count=[" + curDiagram.DiagramObjects.Count + "] and curNumConnectors = " + curNumConnectors + "!!!" );

	return curNumConnectors;

}

