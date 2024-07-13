//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Logging
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExExportElementsJAVA

/*
 * This code has been included from the default Diagram Script template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.
 *
 * Script Name:	DImExExportElementsJAVA
 * Author:		J de Baat
 * Purpose:		Export the information from Elements in the selected Diagram using the BaatScriptLib scripts
 * Date:		13-07-2024
 * 
 */

/*
 * Diagram Script main function
 */
function DImExExportElementsJAVA()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started DImExExportElementsJAVA " );


	// Get and check the global variables
	const validDiagram = IMEXGGetAndCheckDiagram();
	if ( validDiagram ) {
		Session.Output( "Diagram is VALID so proceed processing!" );

		let curResult = ImExExportElements();
		if ( curResult.length > 0 ) {
			LOGError( curResult );
		} else {
			Session.Output( "DImExExportElementsJAVA finished processing!" );
		}

	} else {
		LOGError( "Diagram is NOT VALID!" );
	}

	Session.Output( "======================================= Finished DImExExportElementsJAVA " );
}

DImExExportElementsJAVA();
