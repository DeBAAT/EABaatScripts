//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExExportElements

/*
 * Script Name:	PImExExportElements
 * Author:		J de Baat
 * Purpose:		Export the information from Elements in the selected Package or Diagram using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

/*
 * Project Browser Script main function
 */
function PImExExportElements()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started PImExExportElements at " + _BLOGGetDisplayDate() + "!" );

	// Get and check the global variables
	const validPackageObject = IMEXGGetAndCheckPackageObject();
	if ( validPackageObject ) {
		Session.Output( "PackageObject is VALID so proceed processing!" );

		let curResult = ImExExportElements();
		if ( curResult.length > 0 ) {
			BLOGError( curResult );
		} else {
			Session.Output( "PImExExportElements finished processing!" );
		}

	} else {
		BLOGError( "PackageObject is NOT VALID!" );
	}

	Session.Output( "======================================= Finished PImExExportElements at " + _BLOGGetDisplayDate() + "!" );
}

PImExExportElements();
