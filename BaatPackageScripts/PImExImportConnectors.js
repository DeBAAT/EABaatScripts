//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExImportConnectors

/*
 * Script Name:	PImExImportConnectors
 * Author:		J de Baat
 * Purpose:		Import the information from Connectors into the selected Package or Diagram using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

/*
 * Project Browser Script main function
 */
function PImExImportConnectors()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started PImExImportConnectors at " + _BLOGGetDisplayDate() + "!" );

	// Get and check the global variables
	const validPackageObject = IMEXGGetAndCheckPackageObject();
	if ( validPackageObject ) {
		Session.Output( "PackageObject is VALID so proceed processing!" );

		let curResult = ImExImportConnectors();
		if ( curResult.length > 0 ) {
			BLOGError( curResult );
		} else {
			Session.Output( "PImExImportConnectors finished processing!" );
		}

	} else {
		BLOGError( "PackageObject is NOT VALID!" );
	}

	Session.Output( "======================================= Finished PImExImportConnectors at " + _BLOGGetDisplayDate() + "!" );
}

PImExImportConnectors();
