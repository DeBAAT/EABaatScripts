//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExImportElements

/*
 * Script Name:	PImExImportElements
 * Author:		J de Baat
 * Purpose:		Import the information from Elements into the selected Package or Diagram using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

/*
 * Project Browser Script main function
 */
function PImExImportElements()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started PImExImportElements at " + _BLOGGetDisplayDate() + "!" );

	// Get and check the global variables
	const validPackageObject = IMEXGGetAndCheckPackageObject();
	if ( validPackageObject ) {
		Session.Output( "PackageObject is VALID so proceed processing!" );

		let curResult = ImExImportElements();
		if ( curResult.length > 0 ) {
			BLOGError( curResult );
		} else {
			Session.Output( "PImExImportElements finished processing!" );
		}

	} else {
		BLOGError( "PackageObject is NOT VALID!" );
	}

	Session.Output( "======================================= Finished PImExImportElements at " + _BLOGGetDisplayDate() + "!" );
}

PImExImportElements();
