//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExImportElements

/*
 * Script Name:	DImExImportElements
 * Author:		J de Baat
 * Purpose:		Import the information from Elements into the selected Diagram using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

/*
 * Diagram Script main function
 */
function DImExImportElements()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started DImExImportElements at " + _BLOGGetDisplayDate() + "!" );

	// Get and check the global variables
	const validDiagram = IMEXGGetAndCheckDiagram();
	if ( validDiagram ) {
		Session.Output( "Diagram is VALID so proceed processing!" );

		let curResult = ImExImportElements();
		if ( curResult.length > 0 ) {
			BLOGError( curResult );
		} else {
			Session.Output( "DImExImportElements finished processing!" );
		}

	} else {
		BLOGError( "Diagram is NOT VALID!" );
	}

	Session.Output( "======================================= Finished DImExImportElements at " + _BLOGGetDisplayDate() + "!" );
}

DImExImportElements();
