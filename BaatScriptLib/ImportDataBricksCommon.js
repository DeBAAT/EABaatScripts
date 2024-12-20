//[group=BaatScriptLib]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector

/*
 * Script Name:	ImportDataBricksCommon
 * Author:		J. de Baat
 * Purpose:		BaatScriptLib scripts to assist importing the information from DataBricks into the repository
 * Date:		20-12-2024
 * 
 */

const IDBTaggedValueIDBKey         = "IDBKey";				// Definition of the TV Property used as IDB Key in DataBricks
const IDBTaggedValueIDBConnectorID = "IDBConnectorID";		// Definition of the TV Property used as IDB Key for the Connector ID
const IDBTaggedValueIDBProcessed   = "IDBProcessed";		// Definition of the TV Property used as IDB Processed value in DataBricks
const IDBTaggedValueIDBSource      = "IDBSource";			// Definition of the TV Property used as IDB Key for the Connector Client
const IDBTaggedValueIDBSourceID    = "IDBSourceID";			// Definition of the TV Property used as IDB Key for the Connector Client
const IDBTaggedValueIDBTarget      = "IDBTarget";			// Definition of the TV Property used as IDB Key for the Connector Supplier
const IDBTaggedValueIDBTargetID    = "IDBTargetID";			// Definition of the TV Property used as IDB Key for the Connector Supplier
const IDBTaggedValueIDBStageName   = "Stage";				// Definition of the TV Property used as IDB Processed value in DataBricks

const IDBColumnIDBKey              = strTaggedValuesPrefix + IDBTaggedValueIDBKey;		// Definition of the Column Name used as IDB key for the ImEx Modules
const IDBColumnIDBProcessed        = strTaggedValuesPrefix + IDBTaggedValueIDBProcessed;	// Definition of the Column Name used as IDB key for the ImEx Modules
const IDBColumnIDBSource           = strTaggedValuesPrefix + IDBTaggedValueIDBSource;	// Definition of the Column Name used as IDB key for the ImEx Modules
const IDBColumnIDBSourceID         = strTaggedValuesPrefix + IDBTaggedValueIDBSourceID;	// Definition of the Column Name used as IDB key for the ImEx Modules
const IDBColumnIDBTarget           = strTaggedValuesPrefix + IDBTaggedValueIDBTarget;	// Definition of the Column Name used as IDB key for the ImEx Modules
const IDBColumnIDBTargetID         = strTaggedValuesPrefix + IDBTaggedValueIDBTargetID;	// Definition of the Column Name used as IDB key for the ImEx Modules
const IDBColumnIDBStageName        = strTaggedValuesPrefix + IDBTaggedValueIDBStageName;	// Definition of the Column Name used as IDB key for the ImEx Modules

const IDBKeyPrefix                 = "IDB_";
const IDBKeyElementColumnName      = "Object_ID";			// Definition of the DB Keys to get the Element   IDB key from a TaggedValue
const IDBKeyElementTable           = "t_objectproperties";	// Definition of the DB Keys to get the Element   IDB key from a TaggedValue
const IDBKeyConnectorColumnName    = "ElementID";			// Definition of the DB Keys to get the Connector IDB key from a TaggedValue
const IDBKeyConnectorTable         = "t_connectortag";		// Definition of the DB Keys to get the Connector IDB key from a TaggedValue

const ImExSourceCSV                = "ImExSourceCSV";
const ImExSourceEXCEL              = "ImExSourceEXCEL";

const strTaggedValuesPrefixLength  = strTaggedValuesPrefix.length;

let   strMessage                   = "";
let   strImExSource                = ImExSourceEXCEL;

// Used to create a cache for IDBElementCache
var IDBElementCacheMap             = null;					// : Map   cache for IDBElements
var IDBElementCacheList            = null;					// : Array cache for IDBElements

/**
 * objGlobalIDBPackage needs to have a objGlobalIDBDiagram and visa versa.
 */
var objGlobalIDBPackage as EA.Package;
var objGlobalIDBDiagram as EA.Diagram;

var mapImportColumnTags  = null;			// : Map with tag names found within the column headers

/*
 * Get the TaggedValues columns for the CSV import
 */
function IDBIGetImportColumnTags()
{

	// Only process the CSV columns when ImExSourceCSV and columns not processed yet
	if ( mapImportColumnTags == null ) {
		if ( strImExSource == ImExSourceCSV ) {

			// Create new mapImportColumnTags object to fill
			mapImportColumnTags = new Map();
			var columnName      = "";

			// Process all CSV imported Columns
			let importColumnListLength         = importColumnList.length;
			let strTaggedValuesPrefixLowerCase = strTaggedValuesPrefix.toLowerCase();

			for ( var i = 0 ; i < importColumnListLength ; i++ ) {

				//	If columnName starts with thePrefix then process value without prefix
				columnName = importColumnList[i];
				if ( columnName.substring( 0, strTaggedValuesPrefixLength ).toLowerCase() == strTaggedValuesPrefixLowerCase ) {
					BLOGTrace( " found columnName.substring(" + strTaggedValuesPrefixLength + ")= " + columnName.substring( strTaggedValuesPrefixLength ) + " !" );
					importColumnTag = columnName.substring( strTaggedValuesPrefixLength );
					if ( importColumnTag.length > 0 ) {
						mapImportColumnTags.set( columnName, importColumnTag );
						BLOGTrace( "mapImportColumnTags.set(" + columnName + "," + importColumnTag + ")!!!" );
					}
				}
			
			}
		}
	}

	return mapImportColumnTags;
}

/*
 * GetColumnValueByName from CSV or EXCEL, depending on the strImExSource
 */
function IDBIGetColumnValueByName( theColumnName )
{

	// Find the curColumnValue identified by theTVName, depending on the strImExSource
	// let curColumnName  = strTaggedValuesPrefix + theTVName;
	let curColumnValue = "";
	if ( strImExSource == ImExSourceCSV ) {
		curColumnValue = CSVIGetColumnValueByName( theColumnName );
	} else {
		curColumnValue = EXCELIGetColumnValueByName( theColumnName );
	}

	// Check whether the curColumnValue is found
	if ( ( curColumnValue == undefined ) || ( curColumnValue == null ) || ( curColumnValue == "" ) ) {
		return null;
	}

	return curColumnValue;
}

/*
 * Find theElement with the GUID information as provided
 */
function IDBFindElementByGUID( theColumnGUID )
{

	var curElement     as EA.Element;
	let curElementGUID  = "";

	// Find and process the Element to be identified by theColumnGUID, depending on the strImExSource
	curElementGUID = IDBIGetColumnValueByName( theColumnGUID );
	if ( curElementGUID == null ) {
		// Return null because not found
		BLOGTrace( " curElementGUID not available !!!" );
		return null;
	}
	curElement     = GetElementByGuid( curElementGUID );
	if ( curElement != null ) {

		// Return curElement found
		strMessage = "( " + curElementGUID + " ) found curElement.Name = " + curElement.Name + " for theColumnGUID " + theColumnGUID + "!!!";
		BLOGTrace( strMessage );
		return curElement;
	}

	//  Nothing found so return null
	strMessage = "( " + curElementGUID + " ) did not find any curElement for theColumnGUID " + theColumnGUID + "!!!";
	BLOGTrace( strMessage );
	return null;

}

/*
 * Find theConnector with the GUID information as provided
 */
function IDBFindConnectorByGUID( theColumnGUID )
{

	var curConnector     as EA.Connector;
	let curConnectorGUID  = "";

	// Find and process the Connector to be identified by theColumnGUID, depending on the strImExSource
	curConnectorGUID = IDBIGetColumnValueByName( theColumnGUID );
	if ( curConnectorGUID == null ) {
		// Return null because not found
		BLOGTrace( " curConnectorGUID not available !!!" );
		return null;
	}
	curConnector     = GetConnectorByGuid( curConnectorGUID );
	if ( curConnector != null ) {

		// Return curConnector found
		strMessage = "( " + curConnectorGUID + " ) found curConnector.ConnectorID = " + curConnector.ConnectorID + " for theColumnGUID " + theColumnGUID + "!!!";
		BLOGTrace( strMessage );
		return curConnector;
	}

	//  Nothing found so return null
	strMessage = "( " + curConnectorGUID + " ) did not find any curConnector for theColumnGUID " + theColumnGUID + "!!!";
	BLOGTrace( strMessage );
	return null;

}

/*
 * Get the TaggedValue related to the theTVProperty with value theTVValue from the theTVTable
 */
function IDBGetReferenceIDFromTVProperty( theColumnName, theTVTable, theTVProperty, theTVValue )
{

	// Get the Object_ID registered for theTVProperty with value theTVValue
	let IDBKeyWhereClause   = theTVTable + ".Property = '" + IDBTaggedValueIDBKey + "'"
							+ " and " + theTVTable + ".Value = '" + theTVValue + "'";
	let curReferenceID = IDBGetFieldValueString( theColumnName, theTVTable, IDBKeyWhereClause );
	if ( curReferenceID.length > 0 ) {

		// Return the curReferenceID for curElement found
		strMessage = "( " + theTVValue + " ) found curReferenceID = " + curReferenceID + "!!!";
		BLOGTrace( strMessage );
		return curReferenceID;
	}

	//  Nothing found so return null
	strMessage = "( " + theTVValue + " ) did not find any curReferenceID for theTVProperty " + theTVProperty + "!!!";
	BLOGTrace( strMessage );
	return null;

}

/**
 * Queries the repository database for the first field value whose corresponding row matches the
 * specified WHERE clause.
 *
 * @param[in] columnName (string) The name of the column whose field value will be queried for
 * @param[in] tableName (string) The name of the table that the column resides in
 * @param[in] whereClause (string) The SQL where clause that the query will use to select the
 * appropriate row
 *
 * @return A String representing the requested field value
 */
function IDBGetFieldValueString( columnName /* : String */, tableName /* : String */, whereClause /* : String */ ) /* : String */
{
	var stringValue = "";
	
	// Construct and execute the querySQL and parse the queryResult
	let querySQL               = "SELECT " + columnName + " FROM " + tableName + " WHERE " + whereClause;
	let queryResult            = Repository.SQLQuery( querySQL );
	let queryResultParts       = queryResult.split( columnName );
	let queryResultPartsLength = queryResultParts.length;

	// BLOGTrace("Repository querySQL = " + querySQL + "!!!" );
	// BLOGTrace("Repository found queryResultPartsLength = " + queryResultPartsLength + ", queryResult = " + queryResult + "!!!" );

	// Check the queryResult for the stringValue found
	if ( queryResultPartsLength == 3 )
	{
		if ( queryResultParts[1].startsWith(">") ) {
			stringValue = queryResultParts[1].substring( 1, ( queryResultParts[1].length - 2 ) );
		}
		strMessage = "Repository found queryResultParts[" + 1 + "] = " + queryResultParts[1] + ", stringValue = " + stringValue + "!!!";
		BLOGTrace( strMessage );
	}

	// Free up memory
	queryResult            = null;
	queryResultParts       = null;
	queryResultPartsLength = null;

	return stringValue;
}

/*
 * Find curElement using the TaggedValue in theTVIDBKey Property
 */
function IDBFindElementByTVIDBKey( theTVIDBKey )
{

	var curElementIndex     = -1;
	var curElement         as EA.Element;
	let curColumnIDBKey     = "";
	let curElementIDBKey    = "";
	let curElementObject_ID = "";

	// Find the Element identified by theTVIDBKey
	curColumnIDBKey  = strTaggedValuesPrefix + theTVIDBKey;
	curElementIDBKey = IDBIGetColumnValueByName( curColumnIDBKey );

	// Find and process the element found by curElementIDBKey from the cache
	curElement = IDBFindElementByTVIDBCache( curElementIDBKey );
	if ( curElement != null ) {
		strMessage = "( " + curElementIDBKey + " ) CACHE found curElement.Name = " + curElement.Name + ", ElementID = " + curElement.ElementID + ", ParentID = " + curElement.ParentID + ", PackageID = " + curElement.PackageID + "!!!";
		BLOGTrace( strMessage );
		return curElement;
	}

	// Find the Element identified by theTVIDBKey
	curElementObject_ID = IDBGetReferenceIDFromTVProperty( IDBKeyElementColumnName, IDBKeyElementTable, theTVIDBKey, curElementIDBKey );
	curElement          = GetElementByID( curElementObject_ID );
	if ( curElement != null ) {
		//  Add curElement to mapIDBElementCache for theTVIDBKeyValue
		curElementIndex = IDBElementCacheMap.size;
		IDBElementCacheList[ curElementIndex ] = curElement;
		IDBElementCacheMap.set( curElementIDBKey, curElementIndex );

		// Return curElement found
		strMessage = "( " + curElementIDBKey + " ) found curElement.Name = " + curElement.Name + " for curElementIDBKey = " + curElementIDBKey + " and curElementObject_ID " + curElementObject_ID + "!!!";
		BLOGTrace( strMessage );
		return curElement;
	}

	//  Nothing found so return null
	strMessage = "( " + theTVIDBKey + " ) did not find curElement for curElementIDBKey = " + curElementIDBKey + " and curElementObject_ID " + curElementObject_ID + "!!!";
	BLOGTrace( strMessage );
	return null;

}

/*
 * Find curConnector using the TaggedValue in theTVIDBKey Property
 */
function IDBFindConnectorByTVIDBKey( theTVIDBKey )
{

	var curConnector         as EA.Connector;
	let curColumnIDBKey       = "";
	let curConnectorIDBKey    = "";
	let curConnectorElementID = "";

	// Find the Connector identified by theTVIDBKey
	curColumnIDBKey    = strTaggedValuesPrefix + theTVIDBKey;
	curConnectorIDBKey = IDBIGetColumnValueByName( curColumnIDBKey );

	// Find the Connector identified by theTVIDBKey
	curConnectorElementID  = IDBGetReferenceIDFromTVProperty( IDBKeyConnectorColumnName, IDBKeyConnectorTable, theTVIDBKey, curConnectorIDBKey );
	curConnector           = GetConnectorByID( curConnectorElementID );
	if ( curConnector != null ) {

		// Return curConnector found
		strMessage = "( " + curConnectorIDBKey + " ) found curConnector.ConnectorID = " + curConnector.ConnectorID + " for curConnectorIDBKey = " + curConnectorIDBKey + " and curConnectorElementID " + curConnectorElementID + "!!!";
		BLOGTrace( strMessage );
		return curConnector;
	}

	//  Nothing found so return null
	strMessage = "( " + theTVIDBKey + " ) did not find curConnector for curConnectorIDBKey = " + curConnectorIDBKey + " and curConnectorElementID " + curConnectorElementID + "!!!";
	BLOGTrace( strMessage );
	return null;

}

/*
 * Find curElement using the TaggedValue in theTVIDBKeyValue in the mapIDBElementCache
 */
function IDBFindElementByTVIDBCache( theTVIDBKeyValue )
{

	var curElementIndex  = -1;
	var curElement      as EA.Element;
	curElement           = null;

	// Create the cache if it does not exist yet
	if ( IDBElementCacheMap == null ) {
		IDBElementCacheMap  = new Map();
		IDBElementCacheList = [];
	}

	// Find the Element identified by theTVIDBKeyValue
	if ( IDBElementCacheMap != null && IDBElementCacheMap.has( theTVIDBKeyValue ) ) {
		curElementIndex = IDBElementCacheMap.get( theTVIDBKeyValue );
		curElement      = IDBElementCacheList[ curElementIndex ];
	}

	if ( curElement != null ) {

		// Return curElement found
		strMessage = " found curElement.Name = " + curElement.Name + " for theTVIDBKeyValue = " + theTVIDBKeyValue + "!!!";
		BLOGTrace( strMessage );
		return curElement;
	}

	//  Nothing found so return null
	strMessage = " did not find curElement for theTVIDBKeyValue = " + theTVIDBKeyValue + "!!!";
	BLOGTrace( strMessage );
	return null;

}

/*
 * Find theElement with information as provided
 */
function IDBFindElement( theColumnGUID, theTaggedValueIDBKey )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement      as EA.Element;

	// Find and process the element found by IDBTaggedValueKeyName
	curElement = IDBFindElementByTVIDBKey( theTaggedValueIDBKey );
	if ( curElement != null ) {
		strMessage = "( " + theTaggedValueIDBKey + " ) found curElement.Name = " + curElement.Name + ", ElementID = " + curElement.ElementID + ", ParentID = " + curElement.ParentID + ", PackageID = " + curElement.PackageID + "!!!";
		BLOGTrace( strMessage );
		return curElement;
	}

	// Find and process the Element to be identified by theColumnGUID
	curElement = IDBFindElementByGUID( theColumnGUID );
	if ( curElement != null ) {
		strMessage = "( " + theColumnGUID + " ) found curElement.Name = " + curElement.Name + ", ElementID = " + curElement.ElementID + ", ParentID = " + curElement.ParentID + ", PackageID = " + curElement.PackageID + "!!!";
		BLOGTrace( strMessage );
		return curElement;
	}

	//  Nothing found so return null
	strMessage = "Did not find curElement for theColumnGUID = " + theColumnGUID + " nor theTaggedValueIDBKey " + theTaggedValueIDBKey + "!!!";
	BLOGTrace( strMessage );
	return null;
}

/*
 * Update the TaggedValues of this Element with information as provided
 */
function IDBISetElementTaggedValues( theElement )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement  as EA.Element;
	let columnValue  = "";

	curElement = theElement;

	// Only process a valid curElement
	if ( curElement != null ) {

		// Get the list of TaggedValues from the importColumnList
		IDBIGetImportColumnTags();

		// Process the TaggedValues found
		mapImportColumnTags.forEach(function(value, key) {

				// Get and check theGlobalPathPackage
				let curTaggedValue = IDBIGetColumnValueByName( key );
				if ( curTaggedValue != null ) {
					// If curTaggedValue in import, add it to the curElement
					TVSetElementTaggedValue( curElement, value, curTaggedValue, true );
					BLOGTrace( "mapImportColumnTags(" + value + "," + key + ") PROCESSING for curElement " + curElement.Name + "!!!" );
				}
			});

		// Commit the updated values
		curElement.Update();

	}

}

/*
 * Update the TaggedValues of this Connector with information as provided
 */
function IDBISetConnectorTaggedValues( theConnector )
{

	// Cast theElement to EA.Connector so we get intellisense
	var curConnector  as EA.Connector;
	let columnValue  = "";

	curConnector = theConnector;

	// Only process a valid curElement
	if ( curConnector != null ) {

		// Get the list of TaggedValues from the importColumnList
		IDBIGetImportColumnTags();

		// Process the TaggedValues found
		mapImportColumnTags.forEach(function(value, key) {

				// Get and check theGlobalPathPackage
				let curTaggedValue = IDBIGetColumnValueByName( key );
				if ( curTaggedValue != null ) {
					// If curTaggedValue in import, add it to the curElement
					TVSetElementTaggedValue( curConnector, value, curTaggedValue, true );
					BLOGTrace( "mapImportColumnTags(" + value + "," + key + ") PROCESSING for curConnector " + curConnector.Name + "!!!" );
				}
			});

		// Commit the updated values
		curConnector.Update();

	}

}

/*
 * Find theConnector with information as provided
 */
function IDBFindConnector( theColumnGUID, theTaggedValueIDBKey )
{

	// Cast theConnector to EA.Connector so we get intellisense
	var curConnector as EA.Connector;

	// Find and process the element found by IDBTaggedValueKeyName
	curConnector = IDBFindConnectorByTVIDBKey( theTaggedValueIDBKey );
	if ( curConnector != null ) {
		strMessage = "( " + theTaggedValueIDBKey + " ) found ConnectorID = " + curConnector.ConnectorID + ", ClientID = " + curConnector.ClientID + ", SupplierID = " + curConnector.SupplierID + "!!!";
		BLOGTrace( strMessage );
		return curConnector;
	}

	// Find and process the Connector to be identified by theColumnGUID
	curConnector = IDBFindConnectorByGUID( theColumnGUID );
	if ( curConnector != null ) {
		strMessage = "( " + theColumnGUID + " ) found ConnectorID = " + curConnector.ConnectorID + ", ClientID = " + curConnector.ClientID + ", SupplierID = " + curConnector.SupplierID + "!!!";
		BLOGTrace( strMessage );
		return curConnector;
	}

	//  Nothing found so return null
	strMessage = "Did not find curConnector for theColumnGUID = " + theColumnGUID + " nor theTaggedValueIDBKey " + theTaggedValueIDBKey + "!!!";
	BLOGTrace( strMessage );
	return null;
}

/*
 * Get the requested CollectionObject from theCollectionObjects using theObjectName as parameter
 */
function IDBGetCollectionObjectByName( theCollectionObjects, theObjectName )
{
	// Cast theCollectionObjects to EA.Collection so we get intellisense
	var curCollectionObjects as EA.Collection;
	var curCollectionObject  as EA.Element;

	// Check all Elements in theCollectionObjects whether the requested curElement is already defined
	curCollectionObjects = theCollectionObjects;

	// Loop over curCollectionObjects to find theObjectName
	let curCollectionObjectsCount = curCollectionObjects.Count;
	for ( var i = 0 ; i < curCollectionObjectsCount ; i++ )
	{
		curCollectionObject = curCollectionObjects.GetAt( i );
		if ( curCollectionObject.Name === theObjectName ) {
			strMessage = "Found CollectionObject ( " + curCollectionObject.Name + " ) as part of " + curCollectionObjectsCount + " CollectionObjects!!!";
			BLOGTrace( strMessage );

			// Clean up memory
			curCollectionObjects = null;

			return curCollectionObject;
		}
		// strMessage = " TESTED CollectionObject ( " + curCollectionObject.Name + " ) against theObjectName " + theObjectName + " !!!";
		// BLOGTrace( strMessage );
	}

	// Clean up memory
	curCollectionObjects = null;
	curCollectionObject  = null;

	// theObjectName not found as part of curCollectionObjects
	strMessage = "DID NOT FIND theObjectName ( " + theObjectName + " ) as part of " + curCollectionObjectsCount + " CollectionObjects!!!";
	BLOGTrace( strMessage );
	return null;

}

/*
 * Add a new PackageDiagram if it is not defined yet
 */
function IDBCheckOrAddPackageDiagram( thePackage, thePackageDiagramName )
{

	try {

		// Validate input parameters
		if ( ( thePackage != null ) && ( thePackageDiagramName != "" ) ) {

			var curPackage  as EA.Package;
			var curDiagrams as EA.Collection;
			var curDiagram  as EA.Diagram;

			// Check all Diagrams in thePackage whether the requested thePackageDiagramName already exists
			curPackage  = thePackage;
			curDiagrams = curPackage.Diagrams;
			curDiagram  = curDiagrams.GetByName( thePackageDiagramName );

			// If curDiagram is not found, create a new diagram
			if ( curDiagram == null )
			{
				// BLOGTrace( " addNew because not found: " + thePackageDiagramName );
				curDiagram = curDiagrams.AddNew( thePackageDiagramName, "Logical" );
				curDiagram.Notes = thePackageDiagramName + " created by ImportDataBricksCommon library.";
				curDiagram.Update();

				curDiagrams.Refresh();
				curPackage.Update();

			}

			BLOGTrace( " found " + curDiagram.Name + " as part of PackageID=" + thePackage.PackageID + " !" );
			return curDiagram;
		} else {
			BLOGError( " could NOT add PackageDiagram " + thePackageDiagramName + "!" );
		}

	} catch (err) {
		BLOGError( " catched error " + err.message + "!" );
	}

	return null;

}

/*
 * Add a new subpackage if it is not defined yet
 */
function IDBCheckOrAddSubPackage( thePackage, theSubPackageName )
{

	var curPackagePackage  as EA.Package;
	var curPackagePackages as EA.Collection;
	var curSubPackage      as EA.Package;
	curSubPackage           = null;

	// Check all packages in thePackage whether the requested package already exists
	curPackagePackages = thePackage.Packages;
	curSubPackage      = curPackagePackages.GetByName( theSubPackageName );

	// If curSubPackage is not found, create a new theSubPackageName
	if ( curSubPackage == null )
	{
		strMessage = " addNew because not found: " + theSubPackageName;
		BLOGTrace( strMessage );
		curSubPackage       = curPackagePackages.AddNew( theSubPackageName, "Class" );
		curSubPackage.Notes = theSubPackageName + " created by the ImportDataBricksCommonJAVA script on " + _BLOGGetDisplayDate() + ".";
		curSubPackage.Update();

		curPackagePackages.Refresh();

	}

	// Clean up memory
	curPackagePackage  = null;
	curPackagePackages = null;

	strMessage = "Found " + curSubPackage.Name + " as part of PackageID = " + thePackage.PackageID + " !";
	BLOGTrace( strMessage );
	return curSubPackage;

}

/*
 * Get and check the global variables for an IDB Diagram Script
 */
function IDBGetAndCheckDiagram()
{

	// Prepare some global variables
	objGlobalIDBPackage = null;
	objGlobalIDBDiagram = null;

	// Get a reference to the current diagram
	objGlobalIDBDiagram = Repository.GetCurrentDiagram();

	if ( objGlobalIDBDiagram != null )
	{
		strMessage = "Found Diagram : " + objGlobalIDBDiagram.Name + "!";
		BLOGTrace( strMessage );

		// Get a reference to the parent package of the current diagram
		objGlobalIDBPackage = Repository.GetPackageByID( objGlobalIDBDiagram.PackageID );
		strMessage = "Found " + objGlobalIDBPackage.Name + " as parent package of objGlobalIDBDiagram " + objGlobalIDBDiagram.Name + " !";
		BLOGTrace( strMessage );

	}
	else
	{
		Session.Prompt( "This script requires a diagram to be visible.", promptOK)
	}

	//	Check objGlobalIDBPackage and objGlobalIDBDiagram again
	if ( ( objGlobalIDBPackage == null ) || ( objGlobalIDBDiagram == null ) ) {
		BLOGError( "Both objGlobalIDBPackage AND objGlobalIDBDiagram should be available!" );
		return false;
	}

	return true;

}
