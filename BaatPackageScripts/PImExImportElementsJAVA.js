//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Logging
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExImportElementsJAVA

/*
 * This code has been included from the default Project Browser template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.   
 * 
 * Script Name:	PImExImportElementsJAVA
 * Author:		J de Baat
 * Purpose:		Import the information from Elements into the selected Package or Diagram using the BaatScriptLib scripts
 * Date:		13-07-2024
 * 
 */

/*
 * Project Browser Script main function
 */
function PImExImportElementsJAVA()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started PImExImportElementsJAVA " );

	// Get and check the global variables
	const validPackageObject = IMEXGGetAndCheckPackageObject();
	if ( validPackageObject ) {
		Session.Output( "PackageObject is VALID so proceed processing!" );

		let curResult = ImExImportElements();
		if ( curResult.length > 0 ) {
			LOGError( curResult );
		} else {
			Session.Output( "PImExImportElementsJAVA finished processing!" );
		}

	} else {
		LOGError( "PackageObject is NOT VALID!" );
	}

	Session.Output( "======================================= Finished PImExImportElementsJAVA " );
}

PImExImportElementsJAVA();
