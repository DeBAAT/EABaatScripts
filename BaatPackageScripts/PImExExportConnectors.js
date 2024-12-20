//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExExportConnectors

/*
 * Script Name:	PImExExportConnectors
 * Author:		J de Baat
 * Purpose:		Export the information from Connectors in the selected Package or Diagram using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

/*
 * Project Browser Script main function
 */
function PImExExportConnectors()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started PImExExportConnectors at " + _BLOGGetDisplayDate() + "!" );

	// Get and check the global variables
	const validPackageObject = IMEXGGetAndCheckPackageObject();
	if ( validPackageObject ) {
		Session.Output( "PackageObject is VALID so proceed processing!" );

		let curResult = ImExExportConnectors();
		if ( curResult.length > 0 ) {
			LOGError( curResult );
		} else {
			Session.Output( "PImExExportConnectors finished processing!" );
		}

	} else {
		BLOGError( "PackageObject is NOT VALID!" );
	}

	Session.Output( "======================================= Finished PImExExportConnectors at " + _BLOGGetDisplayDate() + "!" );
}

PImExExportConnectors();
