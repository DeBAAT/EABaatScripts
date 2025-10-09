//[group=BaatScriptLib]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Attribute
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx

/*
 * Script Name:	ImExExportAttributes
 * Author:		J de Baat
 * Purpose:		Export the information from Attributes in the selected Package or Diagram
 * Date:		08-10-2025
 * 
 * Note:	    Open Excel file for writing contents
 */

/*
 * Handle the Excel Application and the WorkSheet to export data to
 */
function ImExExportAttributes( )
{

	// Initialize the EXCEL Export session with the sheet and columns
	let curHandleExcelExportResult = IMEXEHandleExcelExport( strDefaultAttributesSheetName );
	return curHandleExcelExportResult;

}

/*
 * Export the Attributes of the Elements found in this Package
 */
function ImExExportPackageObjects()
{
	let curNumAttributes = 0;

	curNumAttributes = ProcessPackage( objGlobalEAPackage, 0 );
	Session.Output("ImExExportPackageAttributes Processed objGlobalEAPackage( " + objGlobalEAPackage.Name + " ) and newNumAttributes = " + curNumAttributes + "!!!" );

}

/*
 * Export the Attributes of the Elements found in this Diagram
 */
function ImExExportDiagramObjects()
{
	let curNumAttributes = 0;

	curNumAttributes = ProcessDiagram( objGlobalEADiagram, 0 );
	Session.Output("ImExExportDiagramAttributes Processed objGlobalEADiagram( " + objGlobalEADiagram.Name + " ) and newNumAttributes = " + curNumAttributes + "!!!" );

}

/*
 * Initialize the EXCEL Export columns for Attributes
 */
function ImExGetStandardObjectColumns()
{

	// Initialize the EXCEL Export columns for Attributes
	let curExportColumns = IMEXEGetStandardAttributeColumns();
	return curExportColumns;

}

/*
 * Process theElement provided as parameter and its Attributes
 */
function ProcessElementAttributes( theElement, theNumAttributes )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement           as EA.Element;
	var curElementAttribute  as EA.Attribute;
	var curElementAttributes as EA.Collection;
	var curAttributeTag      as EA.TaggedValue;
	var curAttributeTags     as EA.Collection;
	let curNumAttributes;
	let newNumAttributes;
	let curTagColumnArray = [];

	curElement       = theElement;
	curNumAttributes = theNumAttributes;

	if ( curElement == undefined ) {
		return theNumAttributes;
	}

	// Process all curElementAttributes in curElement
	curElementAttributes = curElement.Attributes;
	let curElementAttributesCount = curElementAttributes.Count;
	for ( let i = 0 ; i < curElementAttributesCount ; i++ )
	{
		curElementAttribute = curElementAttributes.GetAt( i );
		// newNumAttributes    = ProcessAttribute( curElementAttribute, curNumAttributes );
		// curNumAttributes    = newNumAttributes;
		// Session.Output("ProcessElementAttributes Processed curElementAttributes(" + i + ") of [" + curElementAttributes.Count + "] and curNumAttributes = " + curNumAttributes + "!!!" );

		// Process all curAttributeTags in curElementAttribute
		curAttributeTags = curElementAttribute.TaggedValues;
		let curAttributeTagsCount = curAttributeTags.Count;
		for ( let j = 0 ; j < curAttributeTagsCount ; j++ )
		{
			curAttributeTag  = curAttributeTags.GetAt( j );
			curTagColumnArray.push( strTaggedValuesPrefix + curAttributeTag.Name );
			// Session.Output("ProcessElementAttributes Processed curAttributeTags(" + j + ") of [" + curAttributeTags.Count + "] and curNumElements = " + curNumElements + "!!!" );
		}

		// Add the TaggedValues to Export columns
		if ( curTagColumnArray.length > 0 ) {
			EXCELEAddExportColumns( curTagColumnArray );
		}

		// Build the curAttributeMap with values to Export
		let curAttributeMap   = IMEXEGetStandardAttributeFieldValues( curElement, curElementAttribute );

		// Export the Attribute of this Element
		if ( curAttributeMap != null ) {
			// let curTaggedValueMap = IMEXEGetAttributeTaggedValues( curAttributeMap, curElementAttribute );
			// EXCELEExportRow( curTaggedValueMap );
			EXCELEExportRow( curAttributeMap );
			curNumAttributes++;
			Session.Output( "ProcessElementAttributes Processed curAttributeMap for ElementID " + curElement.ElementID + " and Attribute.Name " + curElementAttribute.Name + "!" );
		} else {
			BLOGError( "ProcessElementAttributes could NOT get curAttributeMap for ElementID " + curElement.ElementID + " and Attribute.Name " + curElementAttribute.Name + "!" );
		}

		// Clean up memory
		curAttributeMap  = null;
	}

	// Clean up memory
	curElementAttribute  = null;
	curElementAttributes = null;
	curAttributeTag      = null;
	curAttributeTags     = null;
	// curTaggedValueMap    = null;

	// Session.Output("ProcessElementAttributes Processed Element(" + curElement.Name + ") with ObjectType(Type)=" + curElement.Type + ", Attributes.Count=[" + curElement.Attributes.Count + "] and curNumAttributes = " + curNumAttributes + "!!!" );

	return curNumAttributes;
}

/*
 * Process thePackage provided as parameter and its Elements and SubPackages
 */
function ProcessPackage( thePackage, theNumAttributes )
{

	// Cast thePackage to EA.Package so we get intellisense
	var curPackage         as EA.Package;
	var curPackageElements as EA.Collection;
	var curPackagePackages as EA.Collection;
	var curPackageElement  as EA.Element;
	var curPackagePackage  as EA.Package;
	let curNumAttributes;
	let newNumAttributes;

	curPackage       = thePackage;
	curNumAttributes = theNumAttributes;

	// Process all Elements in curPackage
	curPackageElements = curPackage.Elements;
	for ( let i = 0 ; i < curPackageElements.Count ; i++ )
	{
		curPackageElement = curPackageElements.GetAt( i );
		newNumAttributes  = ProcessElementAttributes( curPackageElement, curNumAttributes );
		curNumAttributes  = newNumAttributes;
		// Session.Output("ProcessPackage Processed curPackageElement(" + i + ") of [" + curPackageElements.Count + "] and curNumAttributes = " + curNumAttributes + "!!!" );
	}

	// Clean up memory
	curPackageElement  = null;
	curPackageElements = null;


	// Process all subPackages in curPackage
	curPackagePackages = curPackage.Packages;
	// Session.Output("ProcessPackage Starting recursively for Package(" + curPackage.Name + ") with curPackagePackages.Count=[" + curPackagePackages.Count + "] and curNumAttributes = " + curNumAttributes + "!!!" );
	for ( let i = 0 ; i < curPackagePackages.Count ; i++ )
	{
		curPackagePackage = curPackagePackages.GetAt( i );
		newNumAttributes  = ProcessPackage( curPackagePackage, curNumAttributes );
		curNumAttributes  = newNumAttributes;
		// Session.Output("curPackagePackage(" + curPackagePackage.Name + ") with Elements.Count=[" + curPackagePackage.Elements.Count + "] and Packages.Count=[" + curPackagePackage.Packages.Count + "] !!!" );
	}

	// Clean up memory
	curPackagePackage  = null;
	curPackagePackages = null;

	Session.Output("ProcessPackage Processed Package(" + curPackage.Name + ") with Elements.Count=[" + curPackage.Elements.Count + "], Packages.Count=[" + curPackage.Packages.Count + "] and curNumAttributes = " + curNumAttributes + "!!!" );

	return curNumAttributes;

}

/*
 * Process theDiagram provided as parameter and its Elements
 */
function ProcessDiagram( theDiagram, theNumAttributes )
{

	// Cast theDiagram to EA.Diagram so we get intellisense
	var curDiagram        as EA.Diagram;
	var curDiagramObjects as EA.Collection;
	var curDiagramObject  as EA.DiagramObject;
	var curDiagramElement as EA.Element;
	let curNumAttributes;
	let newNumAttributes;

	curDiagram       = theDiagram;
	curNumAttributes = theNumAttributes;

	// Process all Elements in curDiagram
	curDiagramObjects = curDiagram.DiagramObjects;
	for ( let i = 0 ; i < curDiagramObjects.Count ; i++ )
	{
		// Get the curDiagramDiagramObject from the Collection
		curDiagramObject = curDiagramObjects.GetAt( i );

		// Get the curDiagramElement using the curDiagramObject.ElementID
		curDiagramElement = Repository.GetElementByID( curDiagramObject.ElementID );
		if ( curDiagramElement != null ) {
			newNumAttributes = ProcessElementAttributes( curDiagramElement, curNumAttributes );
			curNumAttributes = newNumAttributes;
			// Session.Output("ProcessDiagram Processed curDiagramElement(" + i + ") of [" + curDiagramObjects.Count + "] and curNumAttributes = " + curNumAttributes + "!!!" );
		}
	}

	// Clean up memory
	curDiagramElement  = null;
	curDiagramElements = null;

	Session.Output("ProcessDiagram Processed Diagram(" + curDiagram.Name + ") with DiagramObjects.Count=[" + curDiagram.DiagramObjects.Count + "] and curNumAttributes = " + curNumAttributes + "!!!" );

	return curNumAttributes;

}

