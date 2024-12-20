//[group=BaatScriptLib]
!INC BaatScriptLib.BaatScript-XML

/**
 * @file BaatScript-Database
 * This script library contains helper functions to assist with querying the underlying database
 * fields of an EA project.
 *
 * Some functions will behave differently according to the value of the script variable DBTYPE.
 * You should ensure that this value reflects the corresponding database repository type that you
 * are currently using.
 *
 * @author	J. de Baat, based on JavaScript-Database by Sparx Systems
 * @date	20-12-2024
 */
var DBTYPE_EAP   = 0;
var DBTYPE_MYSQL = 1;
var DBTYPE       = DBTYPE_EAP;

/**
 * Queries the repository database for the first field value whose corresponding row matches the
 * specified WHERE clause.
 *
 * @param[in] columnName (string) The column name whose field value will be queried for
 * @param[in] table (string) The table that the column resides in
 * @param[in] whereClause (string) The SQL where clause that the query will use to select the
 * appropriate row
 *
 * @return A String representing the requested field value
 */
function DBGetFieldValueString( columnName /* : String */, table /* : String */, whereClause /* : String */ ) /* : String */
{
	var stringValue = "";
	
	// Construct and execute the query
	var sql = "SELECT " + columnName + " FROM " + table + " WHERE " + whereClause;
	var queryResult = Repository.SQLQuery( sql );
	
	if ( queryResult.length > 0 )
	{
		var resultDOM = XMLParseXML( queryResult );
		if ( resultDOM )
			stringValue = XMLGetNodeText( resultDOM, "//EADATA//Dataset_0//Data//Row//" + columnName );
	}

	// Free up memory
	queryResult = null;
	resultDOM   = null;

	return stringValue;
}

/**
 * Queries the repository database for the first field value whose corresponding row matches the
 * specified WHERE clause.
 *
 * @param[in] columnName (string) The column name whose field value will be queried for
 * @param[in] table (string) The table that the column resides in
 * @param[in] whereClause (string) The SQL where clause that the query will use to select the
 * appropriate row
 *
 * @return A Number representing the requested field value
 */
function DBGetFieldValueNumber( columnName /* : String */, table /* : String */, whereClause /* : String */ ) /* : Number */
{
	// Get the field value as a String
	var numberValue = 0;
	var valueAsString = DBGetFieldValueString( columnName, table, whereClause );
	
	// Conver to number
	if ( valueAsString.length > 0 )
		numberValue = new Number( valueAsString );
	
	return numberValue;
}

/**
 * Queries the repository database for all field values whose corresponding row match the specified
 * WHERE clause.
 *
 * @param[in] columnName (string) The column name whose values will be queried for
 * @param[in] table (string) The table that the column resides in
 * @param[in] whereClause (string) The SQL where clause that the query will use to select the
 * appropriate rows
 *
 * @return An array of Strings representing the requested field values
 */
function DBGetFieldValueArrayString( columnName /* : String */, table /* : String */, 
	whereClause /* : String */ ) /* : Array */
{
	var resultArray = new Array();
	
	// Construct and execute the query
	var sql = "SELECT " + columnName + " FROM " + table;
	if ( typeof(whereClause) != "undefined" && whereClause.length > 0 )
		sql += " WHERE " + whereClause;
	
	var queryResult = Repository.SQLQuery( sql );
	if ( queryResult.length > 0 )
	{
		var resultDOM = XMLParseXML( queryResult );
		resultArray = XMLGetNodeTextArray( resultDOM, "//EADATA//Dataset_0//Data//Row//" 
			+ columnName );
	}
	
	return resultArray;
}

/**
 * Returns an escaped copy of the provided String that may be safely included in an SQL query.
 * NOTE: This function automatically adds single quotation marks around the string value.
 *
 * @param[in] originalString The String to escape
 *
 * @return A String representing the SQL escaped version of the provided string
 */
function DBSafeSQLString( originalString /* : String */ ) /* : String */
{
	// Replace single quotation marks with 2x single quotation marks
	var quotationRegEx = new RegExp( "\'", "gm" );
	var modifiedContents = originalString.replace( quotationRegEx, "\'\'" );
	
	return "\'" + modifiedContents + "\'";
}

/**
 * Returns a string representation of the provided Date which may be used in SQL queries
 * NOTE: The output of this function depends on the value of the script variable DBTYPE.
 *
 * @param[in] scriptingDate The Date to the format
 *
 * @return A String representing the SQL formatted version of the provided Date
 */
function DBGetSQLDate( scriptingDate /* : Date */ ) /* : String */
{
	var dateDelimiter = "#";
	
	if ( DBTYPE == DBTYPE_MYSQL )
		dateDelimiter = "\'";
	
	var sqlDate = dateDelimiter;
	sqlDate += scriptingDate.getFullYear() + "-" + (scriptingDate.getMonth() + 1) + "-" + scriptingDate.getDate();
	sqlDate += " " + scriptingDate.getHours() + ":" + scriptingDate.getMinutes() + ":" + scriptingDate.getSeconds();
	sqlDate += dateDelimiter;
	
	return sqlDate;
}

/**
 * Run SQL and return the ResultSet as a JSON object
 *
 * @param[in] sql (string) The SQL SELECT statement to be executed
 *
 * @return A JSON object with three properties: SQL (string), Rows (Array) and Columns (Array).
 * SQL will contain a copy of the sql statement that was passed to this function. Columns will
 * contain an Array of the unique column names returned by the query. Rows will contain an Array
 * of JSON objects, the properties of each object are based on the columns returned by the SQL Query.
 */
function DBSQLQueryToJSON( sql /* : String */ ) /* : Object */
{
	// Create a new JSON object to represent the result set. Object has three properties: SQL (string), Columns (Array) and Rows (Array)
	var resultSet = {
		"SQL" : sql,
		"Columns" : [],
		"Rows" : []
		};
	
	var xml = Repository.SQLQuery(sql);
	var xmlDOM = XMLParseXML(xml);
	var rowCount = 0;

	var xmlRows = xmlDOM.documentElement.selectNodes( "//EADATA//Dataset_0//Data//Row" );
	if (xmlRows != null)
	{
		// Loop each Row node in the xml
		var xmlRow = xmlRows.nextNode();
		while (xmlRow != null)
		{
			// Create a new JSON Object for each row
			var row = {};
			
			// Loop each node which is a child of the current row (i.e. the columns)
			var xmlColumns = xmlRow.childNodes;
			var xmlColumn = xmlColumns.nextNode();
			while (xmlColumn != null)
			{
				// Create a property on the row object with the same name and value as the node found in the xml
				row[xmlColumn.nodeName] = xmlColumn.text;

				// For first row only, add all unique column names to the Columns array
				if (rowCount == 0 && !resultSet.Columns.includes(xmlColumn.nodeName))
					resultSet.Columns.push(xmlColumn.nodeName);
				
				// Next column
				xmlColumn = xmlColumns.nextNode();
			}
			
			// Append new row to ResultSet.Rows
			resultSet.Rows.push(row);
			rowCount++;
			
			// Next row
			xmlRow = xmlRows.nextNode();
		}
	}
	
	return resultSet;
}
