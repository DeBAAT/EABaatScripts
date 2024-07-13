//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Logging
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx

/*
 * This code has been included from the default Project Browser template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.   
 * 
 * Script Name:	ImExExportElementsJAVA
 * Author:		J de Baat
 * Purpose:		Export the information from Elements in the selected Package or Diagram
 * Date:		13-07-2024
 * 
 */

/*
 * Handle the Excel Application and the WorkSheet to export data to
 */
function ImExExportElements( )
{

	// Initialize the EXCEL Export session with the sheet and columns
	let curHandleExcelExportResult = IMEXEHandleExcelExport( strDefaultElementsSheetName );
	return curHandleExcelExportResult;

}

/*
 * Export the Elements found in this Package
 */
function ImExExportPackageObjects()
{
	let curNumElements = 0;

	curNumElements = ProcessPackage( objGlobalEAPackage, 0 );
	Session.Output("ImExExportPackageElements Processed objGlobalEAPackage( " + objGlobalEAPackage.Name + " ) for NumElements = " + curNumElements + "!!!" );

}

/*
 * Export the Elements found in this Diagram
 */
function ImExExportDiagramObjects()
{
	let curNumElements = 0;

	curNumElements = ProcessDiagram( objGlobalEADiagram, 0 );
	Session.Output("ImExExportDiagramElements Processed objGlobalEADiagram( " + objGlobalEADiagram.Name + " ) for NumElements = " + curNumElements + "!!!" );

}

/*
 * Initialize the EXCEL Export columns for Elements
 */
function ImExGetStandardObjectColumns()
{

	// Initialize the EXCEL Export columns for Elements
	let curExportColumns = IMEXEGetStandardElementColumns();
	return curExportColumns;

}

/*
 * Process theTaggedValue provided as parameter
 */
function ProcessTaggedValue( theTaggedValue, theNumElements )
{
	var curTaggedValue as EA.TaggedValue;
	let curNumElements;

	curTaggedValue = theTaggedValue;
	// curNumElements = theNumElements + 1;
	curNumElements = theNumElements;

	// Session.Output("ProcessTaggedValue Processed curTaggedValue(" + curTaggedValue.Name + ") with Value=[" + curTaggedValue.Value + "] and curNumElements = " + curNumElements + "!!!" );

	return curNumElements;

}

/*
 * Process theElement provided as parameter and its TaggedValues
 */
function ProcessElement( theElement, theNumElements )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement     as EA.Element;
	var curElementTag  as EA.TaggedValue;
	var curElementTags as EA.Collection;
	let curNumElements;
	let newNumElements;
	let curTagColumnArray = [];

	curElement     = theElement;
	curNumElements = theNumElements + 1;

	if ( curElement == undefined ) {
		return theNumElements;
	}

	// Process all curElementTags in curElement
	curElementTags = curElement.TaggedValues;
	for ( let i = 0 ; i < curElementTags.Count ; i++ )
	{
		curElementTag  = curElementTags.GetAt( i );
		newNumElements = ProcessTaggedValue( curElementTag, curNumElements );
		curNumElements = newNumElements;
		curTagColumnArray.push( strTaggedValuesPrefix + curElementTag.Name );
		// Session.Output("ProcessElement Processed curElementTags(" + i + ") of [" + curElementTags.Count + "] and curNumElements = " + curNumElements + "!!!" );
	}

	// Add the TaggedValues to Export columns
	if ( curTagColumnArray.length > 0 ) {
		EXCELEAddExportColumns( curTagColumnArray );
	}

	// Build the curValueMap with values to Export
	let curValueMap       = IMEXEGetStandardElementFieldValues( curElement );
	let curTaggedValueMap = IMEXEGetElementTaggedValues( curValueMap, curElement );

	// Export the element
	EXCELEExportRow( curTaggedValueMap );

	// Clean up memory
	curElementTag     = null;
	curElementTags    = null;
	curValueMap       = null;
	curTaggedValueMap = null;

	// Session.Output("ProcessElement Processed Element(" + curElement.Name + ") with ObjectType(Type)=" + curElement.Type + ", TaggedValues.Count=[" + curElement.TaggedValues.Count + "] and curNumElements = " + curNumElements + "!!!" );

	return curNumElements;
}

/*
 * Process thePackage provided as parameter and its Elements and SubPackages
 */
function ProcessPackage( thePackage, theNumElements )
{

	// Cast thePackage to EA.Package so we get intellisense
	var curPackage         as EA.Package;
	var curPackageElements as EA.Collection;
	var curPackagePackages as EA.Collection;
	var curPackageElement  as EA.Element;
	var curPackagePackage  as EA.Package;
	let curNumElements;
	let newNumElements;

	curPackage     = thePackage;
	curNumElements = theNumElements;

	// Process all Elements in curPackage
	curPackageElements = curPackage.Elements;
	for ( let i = 0 ; i < curPackageElements.Count ; i++ )
	{
		curPackageElement = curPackageElements.GetAt( i );
		newNumElements    = ProcessElement( curPackageElement, curNumElements );
		curNumElements    = newNumElements;
		// Session.Output("ProcessPackage Processed curPackageElement(" + i + ") of [" + curPackageElements.Count + "] and curNumElements = " + curNumElements + "!!!" );
	}

	// Clean up memory
	curPackageElement  = null;
	curPackageElements = null;


	// Recursively Process all subPackages in curPackage
	curPackagePackages = curPackage.Packages;
	// Session.Output("ProcessPackage Starting recursively for Package(" + curPackage.Name + ") with curPackagePackages.Count=[" + curPackagePackages.Count + "] and curNumElements = " + curNumElements + "!!!" );
	for ( let i = 0 ; i < curPackagePackages.Count ; i++ )
	{
		curPackagePackage = curPackagePackages.GetAt( i );
		newNumElements    = ProcessPackage( curPackagePackage, curNumElements );
		curNumElements    = newNumElements;
		// Session.Output("curPackagePackage(" + curPackagePackage.Name + ") with Elements.Count=[" + curPackagePackage.Elements.Count + "] and Packages.Count=[" + curPackagePackage.Packages.Count + "] !!!" );
	}

	// Clean up memory
	curPackagePackage  = null;
	curPackagePackages = null;

	Session.Output("ProcessPackage Processed Package(" + curPackage.Name + ") with Elements.Count=[" + curPackage.Elements.Count + "], Packages.Count=[" + curPackage.Packages.Count + "] for NumElements = " + curNumElements + "!!!" );

	return curNumElements;

}

/*
 * Process theDiagram provided as parameter and its Elements
 */
function ProcessDiagram( theDiagram, theNumElements )
{

	// Cast theDiagram to EA.Diagram so we get intellisense
	var curDiagram        as EA.Diagram;
	var curDiagramObjects as EA.Collection;
	var curDiagramObject  as EA.DiagramObject;
	var curDiagramElement as EA.Element;
	let curNumElements;
	let newNumElements;

	curDiagram     = theDiagram;
	curNumElements = theNumElements;

	// Process all Elements in curDiagram
	curDiagramObjects = curDiagram.DiagramObjects;
	for ( let i = 0 ; i < curDiagramObjects.Count ; i++ )
	{
		// Get the curDiagramDiagramObject from the Collection
		curDiagramObject = curDiagramObjects.GetAt( i );

		// Get the curDiagramElement using the curDiagramObject.ElementID
		curDiagramElement = Repository.GetElementByID( curDiagramObject.ElementID );
		if ( curDiagramElement != null ) {
			newNumElements = ProcessElement( curDiagramElement, curNumElements );
			curNumElements = newNumElements;
			// Session.Output("ProcessDiagram Processed curDiagramElement(" + i + ") of [" + curDiagramObjects.Count + "] for NumElements = " + curNumElements + "!!!" );
		}
	}

	// Clean up memory
	curDiagramObjects = null;
	curDiagramObject  = null;
	curDiagramElement = null;

	Session.Output("ProcessDiagram Processed Diagram(" + curDiagram.Name + ") with DiagramObjects.Count=[" + curDiagram.DiagramObjects.Count + "] for NumElements = " + curNumElements + "!!!" );

	return curNumElements;

}
