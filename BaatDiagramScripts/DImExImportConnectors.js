//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExImportConnectors

/*
 * Script Name:	DImExImportConnectors
 * Author:		J de Baat
 * Purpose:		Import the information from Connectors into the selected Diagram using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

/*
 * Diagram Script main function
 */
function DImExImportConnectors()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started DImExImportConnectors at " + _BLOGGetDisplayDate() + "!" );

	// Get and check the global variables
	const validDiagram = IMEXGGetAndCheckDiagram();
	if ( validDiagram ) {
		Session.Output( "Diagram is VALID so proceed processing!" );

		let curResult = ImExImportConnectors();
		if ( curResult.length > 0 ) {
			BLOGError( curResult );
		} else {
			Session.Output( "DImExImportConnectors finished processing!" );
		}

	} else {
		BLOGError( "Diagram is NOT VALID!" );
	}

	Session.Output( "======================================= Finished DImExImportConnectors at " + _BLOGGetDisplayDate() + "!" );
}

DImExImportConnectors();
