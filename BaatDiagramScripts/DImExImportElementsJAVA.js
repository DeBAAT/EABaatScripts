//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Logging
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExImportElementsJAVA

/*
 * This code has been included from the default Diagram Script template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.
 *
 * Script Name:	DImExImportElementsJAVA
 * Author:		J de Baat
 * Purpose:		Import the information from Elements into the selected Diagram using the BaatScriptLib scripts
 * Date:		13-07-2024
 * 
 */

/*
 * Diagram Script main function
 */
function DImExImportElementsJAVA()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started DImExImportElementsJAVA " );

	// Get and check the global variables
	const validDiagram = IMEXGGetAndCheckDiagram();
	if ( validDiagram ) {
		Session.Output( "Diagram is VALID so proceed processing!" );

		let curResult = ImExImportElements();
		if ( curResult.length > 0 ) {
			LOGError( curResult );
		} else {
			Session.Output( "DImExImportElementsJAVA finished processing!" );
		}

	} else {
		LOGError( "Diagram is NOT VALID!" );
	}

	Session.Output( "======================================= Finished DImExImportElementsJAVA " );
}

DImExImportElementsJAVA();
