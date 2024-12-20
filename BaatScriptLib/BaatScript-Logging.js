//[group=BaatScriptLib]
/**
 * @file BaatScript-Logging
 * This script library contains helper functions to assist with logging. Log messages can be 
 * submitted to the various functions in this module, and will be filtered according to the value
 * of BLOGLEVEL.
 *
 * Valid values for BLOGLEVEL are:
 * - BLOGLEVEL_ERROR
 * - BLOGLEVEL_WARNING
 * - BLOGLEVEL_INFO
 * - BLOGLEVEL_DEBUG
 * - BLOGLEVEL_TRACE
 *
 * You can change the log level at any time during execution by setting the BLOGLEVEL variable to the
 * desired value.
 *
 * Functions provided by this module are identified by the prefix BLOG
 *
 * @author	J. de Baat, based on JavaScript-Logging by Sparx Systems
 * @date	20-12-2024
 */

// BLOGLEVEL values
var BLOGLEVEL_ERROR   = 0;
var BLOGLEVEL_WARNING = 1;
var BLOGLEVEL_INFO    = 2;
var BLOGLEVEL_DEBUG   = 3;
var BLOGLEVEL_TRACE   = 4;

// The level to log at
var BLOGLEVEL = BLOGLEVEL_INFO;

/**
 * Logs a message at the ERROR level. The message will be displayed if BLOGLEVEL is set to 
 * BLOGLEVEL_ERROR or above.
 *
 * @param[in] message (String) The message to log
 */
function BLOGError( message /* : String */ ) /* : void */
{
	if ( BLOGLEVEL >= BLOGLEVEL_ERROR ) {
		Session.Output( _BLOGGetDisplayDate() + " [ERROR] " + BLOGError.caller.name + ":: " + message );
	}
}

/**
 * Logs a message at the INFO level. The message will be displayed if BLOGLEVEL is set to 
 * BLOGLEVEL_INFO or above.
 *
 * @param[in] message (String) The message to log
 */
function BLOGInfo( message /* : String */ ) /* : void */
{
	if ( BLOGLEVEL >= BLOGLEVEL_INFO ) {
		Session.Output( _BLOGGetDisplayDate() + " [INFO] " + BLOGInfo.caller.name + ":: " + message );
	}
}

/**
 * Logs a message at the WARNING level. The message will be displayed if BLOGLEVEL is set to 
 * BLOGLEVEL_WARNING or above.
 *
 * @param[in] message (String) The message to log
 */
function BLOGWarning( message /* : String */ ) /* : void */
{
	if ( BLOGLEVEL >= BLOGLEVEL_WARNING ) {
		Session.Output( _BLOGGetDisplayDate() + " [WARNING] " + BLOGWarning.caller.name + ":: " + message );
	}
}

/**
 * Logs a message at the DEBUG level. The message will be displayed if BLOGLEVEL is set to 
 * BLOGLEVEL_DEBUG or above.
 *
 * @param[in] message (String) The message to log
 */
function BLOGDebug( message /* : String */ ) /* : void */
{
	if ( BLOGLEVEL >= BLOGLEVEL_DEBUG ) {
		Session.Output( _BLOGGetDisplayDate() + " [DEBUG] " + BLOGDebug.caller.name + ":: " + message );
	}
}

/**
 * Logs a message at the TRACE level. The message will be displayed if BLOGLEVEL is set to 
 * BLOGLEVEL_TRACE or above.
 *
 * @param[in] message (String) The message to log
 */
function BLOGTrace( message /* : String */ ) /* : void */
{
	if ( BLOGLEVEL >= BLOGLEVEL_TRACE ) {
		Session.Output( _BLOGGetDisplayDate() + " [TRACE] " + BLOGTrace.caller.name + ":: " + message );
	}
}

/**
 * Returns the current date/time in a format suitable for logging.
 *
 * @return A String representing the current date/time
 */
function _BLOGGetDisplayDate() /* : String */
{
	var now = new Date();
	
	// Pad hour value
	var hours = now.getHours();
	if ( hours < 10 )
		hours = "0" + hours;
	
	// Pad minute value
	var minutes = now.getMinutes();
	if ( minutes < 10 )
		minutes = "0" + minutes;
	
	// Pad second value
	var seconds = now.getSeconds();
	if ( seconds < 10 )
		seconds = "0" + seconds;
	
	var displayDate = now.getFullYear() + "-" + (now.getMonth() + 1) + "-" + now.getDate();
	displayDate += " " + hours + ":" + minutes + ":" + seconds;

	return displayDate;
}