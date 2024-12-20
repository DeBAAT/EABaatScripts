//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExExportElements

/*
 * Script Name:	DImExExportElements
 * Author:		J de Baat
 * Purpose:		Export the information from Elements in the selected Diagram using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

/*
 * Diagram Script main function
 */
function DImExExportElements()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started DImExExportElements at " + _BLOGGetDisplayDate() + "!" );


	// Get and check the global variables
	const validDiagram = IMEXGGetAndCheckDiagram();
	if ( validDiagram ) {
		Session.Output( "Diagram is VALID so proceed processing!" );

		let curResult = ImExExportElements();
		if ( curResult.length > 0 ) {
			BLOGError( curResult );
		} else {
			Session.Output( "DImExExportElements finished processing!" );
		}

	} else {
		BLOGError( "Diagram is NOT VALID!" );
	}

	Session.Output( "======================================= Finished DImExExportElements at " + _BLOGGetDisplayDate() + "!" );
}

DImExExportElements();
