//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Logging
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExExportElementsJAVA

/*
 * This code has been included from the default Project Browser template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.   
 * 
 * Script Name:	PImExExportElementsJAVA
 * Author:		J de Baat
 * Purpose:		Export the information from Elements in the selected Package or Diagram using the BaatScriptLib scripts
 * Date:		13-07-2024
 * 
 */

/*
 * Project Browser Script main function
 */
function PImExExportElementsJAVA()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started PImExExportElementsJAVA " );

	// Get and check the global variables
	const validPackageObject = IMEXGGetAndCheckPackageObject();
	if ( validPackageObject ) {
		Session.Output( "PackageObject is VALID so proceed processing!" );

		let curResult = ImExExportElements();
		if ( curResult.length > 0 ) {
			LOGError( curResult );
		} else {
			Session.Output( "PImExExportElementsJAVA finished processing!" );
		}

	} else {
		LOGError( "PackageObject is NOT VALID!" );
	}

	Session.Output( "======================================= Finished PImExExportElementsJAVA " );
}

PImExExportElementsJAVA();
