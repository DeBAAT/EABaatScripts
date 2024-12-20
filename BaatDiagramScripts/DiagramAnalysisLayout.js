//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.DiagramAnalysisLayoutCommon

/*
 * Script Name:	DiagramAnalysisLayout
 * Author:		J de Baat
 * Purpose:		Dynamically draw elements on the diagram as indicated by the TaggedValues defined in the selected element
 * Date:		20-12-2024
 *
 * This script takes a number of selected elements on a diagram to start the analysis.
 * For each selected element on a diagram
 *   If the Keywords property of the element contains the DALKeywordsTag
 *     Collect all referenced elements to a set for the first level
 *     Process the current level for all elements in the set:
 *       For each referenced element in the set
 *         If element not yet drawn on the diagram
 *           Draw the element on the indicated location of the diagram
 *           Collect all referenced elements to a set for the next level
 *       Recursively process the set of elements found
 *
 */

/*
 * Diagram Script main function
 */
function DiagramAnalysisLayout()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started DiagramAnalysisLayout at " + _BLOGGetDisplayDate() + "!" );

	// Get a reference to objGlobalDALDiagram
	objGlobalDALDiagram = Repository.GetCurrentDiagram();

	if ( objGlobalDALDiagram != null )
	{
		// Get a reference to any selected objects
		var selectedElement as EA.Element;
		var selectedObjects as EA.Collection;
		selectedObjects      = objGlobalDALDiagram.SelectedObjects;

		try
		{

			if ( selectedObjects.Count > 0 )
			{
				BLOGDebug("DiagramAnalysisLayout: Selected selectedObjects.Count: " + selectedObjects.Count );

				// One or more diagram objects are selected
				let selectedObjectsCount = selectedObjects.Count;
				for ( var i = 0 ; i < selectedObjectsCount ; i++ )
				{
					// Process the currentDiagramElement
					var currentDiagramElement as EA.Element;
					var currentElement        as EA.Element;
					currentDiagramElement      = selectedObjects.GetAt( i );
					currentElement             = Repository.GetElementByID( currentDiagramElement.ElementID );

					// Process the currentElement for DiagramAnalysisLayout
					DALProcessElement( currentElement );
				}

				// Reload objGlobalDALDiagram when all processing is done
				Repository.ReloadDiagram( objGlobalDALDiagram.DiagramID );

			}
			else
			{
				// Nothing is selected
				BLOGError( "This script requires at least one element to be selected!" );
			}
		}
		catch (catch_err)
		{
			BLOGError("DiagramAnalysisLayout found Error: " + catch_err.message + "!!!" );
		}

		Session.Output( "======================================= Finished DiagramAnalysisLayout at " + _BLOGGetDisplayDate() + "!" );
	}
	else
	{
		Session.Prompt( "This script requires a diagram to be visible.", promptOK);
	}
}

DiagramAnalysisLayout();
