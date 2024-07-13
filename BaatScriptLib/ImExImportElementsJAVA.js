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
 * Script Name:	ImExImportElementsJAVA
 * Author:		J de Baat
 * Purpose:		Import the information from Elements into the selected Package
 * Date:		13-07-2024
 * 
 */

/*
 * Handle the ExcelImport for importing Elements
 */
function ImExImportElements( )
{

	// Start the ExcelImport for this Import session
	// Session.Output("ImExImportElements started Excel.Application !" );
	let strExcelImportResult = IMEXIHandleExcelImport( strDefaultElementsSheetName );

	// Return the result found
	return strExcelImportResult;

}

/*
 * Process the Imported Elements found
 */
function OnExcelRowImported( theRow )
{

	let actionResult = "";
	let importAction = "";
	let importRow    = 0;

	if ( excelImportCurrentRow[0] != null ) {
		importAction = excelImportCurrentRow[0].toLowerCase();
	}
	importRow = theRow;
	// Session.Output("OnExcelRowImported excelImportCurrentRow(" + importRow + ").length = " + excelImportCurrentRow.length + ", Action = " + excelImportCurrentRow[0] + "!!!" );

	// Process importAction requested
	switch( importAction )
	{
		case "create":
		{
			 actionResult = createElement();
			 break;
		}
		case "update":
		{
			 actionResult = updateElement();
			 break;
		}
		case "delete":
		{
			 actionResult = deleteElement();
			 break;
		}
		case "":
		{
			 actionResult = "OnExcelRowImported found empty importAction!";
			 break;
		}
		default:
		{
			// Error message
			LOGError( "OnExcelRowImported CANNOT process importAction " + importAction + "!"  );
		}
	}

}

/*
 * Create theElement with information as provided
 */
function createElement()
{

	var curElement as EA.Element;

	// Find and process the element found by curElementGUID
	let curElementGUID = EXCELIGetColumnValueByName("CLASSGUID");
	curElement = GetElementByGuid( curElementGUID );
	if ( curElement != null ) {
		return "createElement( " + curElementGUID + " ) found curElement.Name = " + curElement.Name + " so skip creation to prevent duplicates!!!";
	} else {

		// createElement could NOT find curElement so create new one
		let curElementName = EXCELIGetColumnValueByName("Name");
		let curElementType = EXCELIGetColumnValueByName("CLASSTYPE");
		// Session.Output("createElement( " + curElementGUID + " ) could NOT find curElement so create new one with Name " + curElementName + "!!!" );
		curElement = objGlobalEAPackage.Elements.AddNew( curElementName, curElementType );

		// createElement created new curElement so update it using the other values found
		if ( curElement != null ) {

			// Process the updates for curElement found
			let curResult = updateElementProperties( curElement );
			// Session.Output("createElement( " + curElementGUID + " ) updated curElement.Name to " + curElement.Name + ", curResult= " + curResult + "!!!" );

			// Commit the changes to the repository
			objGlobalEAPackage.Update();
			objGlobalEAPackage.Elements.Refresh();

		} else {

			return "createElement( " + curElementGUID + " ) could NOT create curElement so NOT updated!!!";
		}
	}

	return "";
}

/*
 * Update theElement with information as provided
 */
function updateElement()
{

	var curElement as EA.Element;

	// Find and process the element found by curElementGUID
	let curElementGUID = EXCELIGetColumnValueByName("CLASSGUID");
	curElement = GetElementByGuid( curElementGUID );
	if ( curElement != null ) {

		// Process the updates for curElement found
		// Session.Output("updateElement( " + curElementGUID + " ) found curElement.Name = " + curElement.Name + ", Visibility = " + curElement.Visibility + "!!!" );
		let curResult = updateElementProperties( curElement );
		// Session.Output("updateElement( " + curElementGUID + " ) updated curElement.Name to " + curElement.Name + ", curResult= " + curResult + "!!!" );

	} else {

		return "updateElement( " + curElementGUID + " ) could NOT find curElement so NOT updated!!!";
	}

	return "";
}

/*
 * Update theElement with information as provided in the fields
 */
function updateElementProperties( theElement )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement     as EA.Element;
	var curElementTag  as EA.TaggedValue;
	var curElementTags as EA.Collection;

	curElement            = theElement;

	// Process theElement
	if ( curElement != null ) {

		curElementTags   = curElement.TaggedValues;

		// Procees StandardElementFieldValues
		// Session.Output("updateElementProperties( " + curElement.ElementGUID + " ) found curElement.Name = " + curElement.Name + ", Visibility = " + curElement.Visibility + "!!!" );
		IMEXISetStandardElementFieldValues( curElement );
		// Session.Output("updateElementProperties( " + curElement.ElementGUID + " ) updated curElement.Name to " + curElement.Name + "!!!" );

		// Procees StandardElementTaggedValues
		// Session.Output("updateElementProperties( " + curElement.ElementGUID + " ) found curElement.Name = " + curElement.Name + ", Visibility = " + curElement.Visibility + "!!!" );
		IMEXISetElementTaggedValues( curElement );
		// Session.Output("updateElementProperties( " + curElement.ElementGUID + " ) updated curElement.Name to " + curElement.Name + "!!!" );

		// Commit the changes to the repository
		objGlobalEAPackage.Update();
		objGlobalEAPackage.Elements.Refresh();

	} else {

		return "updateElementProperties() could NOT find curElement so NOT updated!!!";
	}

	return "";
}

/*
 * Delete theElement with information as provided
 */
function deleteElement()
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement as EA.Element;

	// Find and process the element found by curElementGUID
	let curElementGUID = EXCELIGetColumnValueByName("CLASSGUID");
	curElement = GetElementByGuid( curElementGUID );
	if ( curElement != null ) {
		// Session.Output("deleteElement( " + curElementGUID + " ) found curElement.Name = " + curElement.Name + ", ElementID = " + curElement.ElementID + ", ParentID = " + curElement.ParentID + ", PackageID = " + curElement.PackageID + "!!!" );
		// Delete the element as part of the curElement.ParentID
		var curParentElement as EA.Element;
		var curTempElement   as EA.Element;
		curParentElement = Repository.GetElementByID( curElement.ParentID );
		if ( curParentElement != null ) {
			// Find the index in the curParentElement.Elements for the curElement to delete
			for ( let i = 0 ; i < curParentElement.Elements.Count ; i++ ) {
				curTempElement = curParentElement.Elements.GetAt( i );
				if ( curTempElement.ElementID == curElement.ElementID ) {
					curParentElement.Elements.DeleteAt( i, false );
					// Session.Output("deleteElement deleted curParentElement(" + i + ") where curElement.ElementID = " + curElement.ElementID + "!!!" );
					break; // Stop processing the rest of the Elements in the for loop
				}
			}
			curParentElement.Elements.Refresh();
		} else {
			// Delete the element as part of the curElement.PackageID
			var curParentPackage as EA.Package;
			curParentPackage = Repository.GetPackageByID( curElement.PackageID );
			if ( curParentPackage != null ) {
				// Find the index in the curParentPackage.Elements for the curElement to delete
				for ( let i = 0 ; i < curParentPackage.Elements.Count ; i++ ) {
					curTempElement = curParentPackage.Elements.GetAt( i );
					if ( curTempElement.ElementID == curElement.ElementID ) {
						curParentPackage.Elements.DeleteAt( i, false );
						// Session.Output("deleteElement deleted curParentPackage(" + i + ") where curElement.ElementID = " + curElement.ElementID + "!!!" );
						break; // Stop processing the rest of the Elements in the for loop
					}
				}
				curParentPackage.Elements.Refresh();
			} else {
				return "deleteElement( " + curElementGUID + " ) could NOT find curElement in ParentID nor PackageID!!!";
			}
		}
	} else {
		return "deleteElement( " + curElementGUID + " ) could NOT find curElement so NOT deleted!!!";
	}

	return "";
}
