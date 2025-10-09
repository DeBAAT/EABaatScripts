//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Attribute
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImExExportAttributes

/*
 * Script Name:	PImExExportAttributes
 * Author:		J de Baat
 * Purpose:		Export the information from Attributes in the selected Package or Diagram using the BaatScriptLib scripts
 * Date:		08-10-2025
 * 
 */

/*
 * Project Browser Script main function
 */
function PImExExportAttributes()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started PImExExportAttributes at " + _BLOGGetDisplayDate() + "!" );

	// Get and check the global variables
	const validPackageObject = IMEXGGetAndCheckPackageObject();
	if ( validPackageObject ) {
		Session.Output( "PackageObject is VALID so proceed processing!" );

		let curResult = ImExExportAttributes();
		if ( curResult.length > 0 ) {
			LOGError( curResult );
		} else {
			Session.Output( "PImExExportAttributes finished processing!" );
		}

	} else {
		BLOGError( "PackageObject is NOT VALID!" );
	}

	Session.Output( "======================================= Finished PImExExportAttributes at " + _BLOGGetDisplayDate() + "!" );
}

PImExExportAttributes();
