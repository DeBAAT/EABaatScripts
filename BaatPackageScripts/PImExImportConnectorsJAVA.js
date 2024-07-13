//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Logging
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExImportConnectorsJAVA

/*
 * This code has been included from the default Project Browser template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.   
 * 
 * Script Name:	PImExImportConnectorsJAVA
 * Author:		J de Baat
 * Purpose:		Import the information from Connectors into the selected Package or Diagram using the BaatScriptLib scripts
 * Date:		13-07-2024
 * 
 */

/*
 * Project Browser Script main function
 */
function PImExImportConnectorsJAVA()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started PImExImportConnectorsJAVA " );

	// Get and check the global variables
	const validPackageObject = IMEXGGetAndCheckPackageObject();
	if ( validPackageObject ) {
		Session.Output( "PackageObject is VALID so proceed processing!" );

		let curResult = ImExImportConnectors();
		if ( curResult.length > 0 ) {
			LOGError( curResult );
		} else {
			Session.Output( "PImExImportConnectorsJAVA finished processing!" );
		}

	} else {
		LOGError( "PackageObject is NOT VALID!" );
	}

	Session.Output( "======================================= Finished PImExImportConnectorsJAVA " );
}

PImExImportConnectorsJAVA();
