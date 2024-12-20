//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExExportConnectors

/*
 * Script Name:	DImExExportConnectors
 * Author:		J de Baat
 * Purpose:		Export the information from Connectors in the selected Diagram using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

/*
 * Diagram Script main function
 */
function DImExExportConnectors()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started DImExExportConnectors at " + _BLOGGetDisplayDate() + "!" );

	// Get and check the global variables
	const validDiagram = IMEXGGetAndCheckDiagram();
	if ( validDiagram ) {
		Session.Output( "Diagram is VALID so proceed processing!" );

		let curResult = ImExExportConnectors();
		if ( curResult.length > 0 ) {
			BLOGError( curResult );
		} else {
			Session.Output( "DImExExportConnectors finished processing!" );
		}

	} else {
		BLOGError( "Diagram is NOT VALID!" );
	}

	Session.Output( "======================================= Finished DImExExportConnectors at " + _BLOGGetDisplayDate() + "!" );
}

DImExExportConnectors();
