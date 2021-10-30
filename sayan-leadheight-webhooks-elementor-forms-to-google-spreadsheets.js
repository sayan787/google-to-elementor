/*
 * In order to enable this script, follow these steps:
   ***   Open the Google Sheet you want to use,
    **   From the Tools menu select "Script Editor"
     *   Paste this script into code editor and hit Save.
     *   
   ***   Then click: 
    **     "Publish" & select "Deploy as web app...
   ***   Select: 
    **      Execute the app as: Me (youremail@gmail.com)
    **      Who has access to the app: Anyone, even anonymous
   *                           - Note, depending on your Google Apps instance, this option may not be available. You will need to contact your Google Apps administrator, or else use a Gmail account. 
   *
   ***    Now click Deploy. You may be asked to review permissions now.
   *                           - It will probably give you a warning, click "Advanced" on the bottom left and continue.
   *      
   ***    The URL that you get will be the webhook that you can use in your elementor form.
   *                            - You can test this webhook in your browser first by pasting it. 
   *                              It will say: "Yepp this is the webhook URL, request received".
   *         
   ***    EMAIL NOTIFICATIONS: 
            *  By default, email notifications are turned off in this script.
            *  To turn them on:
              *  on line 37 Change "false to "true"
              *  on line 40 replace "Change_to_your_Email" with, well, whatever your email is, but leave the "quotes"
              *  re-save the script. collect leades, make that $$$
 */    

function doGet(e) {}
function doPost(e) {
var o = e.parameter;
SpreadsheetApp.getActiveSheet().appendRow([ o.name, o.email, o.message, e ]);
}


// Change to true to enable email notifications
var emailNotification = false;

// Enter your email address below (keep the quotation marks!) 
var emailAddress = "Change_to_your_Email";

/**
 * Google app-script to utilise Elementor Pro Froms webhook
 * For Usage see: https://github.com/pojome/elementor/issues/5894
 * Origionally found: https://gist.github.com/bainternet/4b539b00a4bd7490ac3809d7ff86bd14
 * by bainternet
 * Minor tweaks to the directions by AvlSEONinja
 */
 
 
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DO NOT EDIT THESE NEXT PARAMS
var isNewSheet = false;
var recivedData = [];

/**
 * this is a function that fires when the webapp receives a GET request
 * Not used but required.
 */
function doGet( e ) {
	return HtmlService.createHtmlOutput( "Yepp this is the webhook URL, request received" );
}

// Webhook Receiver - triggered with form webhook to pusblished App URL.
function doPost( e ) {
	var params = JSON.stringify(e.parameter);
	params = JSON.parse(params);
	insertToSheet(params);

	// HTTP Response
	return HtmlService.createHtmlOutput( "post request received" );
}

// Flattens a nested object for easier use with a spreadsheet
function flattenObject( ob ) {
	var toReturn = {};
	for ( var i in ob ) {
		if ( ! ob.hasOwnProperty( i ) ) continue;
		if ( ( typeof ob[ i ] ) == 'object' ) {
			var flatObject = flattenObject( ob[ i ] );
			for ( var x in flatObject ) {
				if ( ! flatObject.hasOwnProperty( x ) ) continue;
				toReturn[ i + '.' + x ] = flatObject[ x ];
			}
		} else {
			toReturn[ i ] = ob[ i ];
		}
	}
	return toReturn;
}

// normalize headers
function getHeaders( formSheet, keys ) {
	var headers = [];
  
	// retrieve existing headers
    if ( ! isNewSheet ) {
	  headers = formSheet.getRange( 1, 1, 1, formSheet.getLastColumn() ).getValues()[0];
    }

	// add any additional headers
	var newHeaders = [];
	newHeaders = keys.filter( function( k ) {
		return headers.indexOf( k ) > -1 ? false : k;
	} );

	newHeaders.forEach( function( h ) {
		headers.push( h );
	} );
	return headers;
}

// normalize values
function getValues( headers, flat ) {
	var values = [];
	// push values based on headers
	headers.forEach( function( h ){
		values.push( flat[ h ] );
	});
	return values;
}

// Insert headers
function setHeaders( sheet, values ) {
	var headerRow = sheet.getRange( 1, 1, 1, values.length )
	headerRow.setValues( [ values ] );
	headerRow.setFontWeight( "bold" ).setHorizontalAlignment( "center" );
}

// Insert Data into Sheet
function setValues( sheet, values ) {
	var lastRow = Math.max( sheet.getLastRow(),1 );
	sheet.insertRowAfter( lastRow );
	sheet.getRange( lastRow + 1, 1, 1, values.length ).setValues( [ values ] ).setFontWeight( "normal" ).setHorizontalAlignment( "center" );
}

// Find or create sheet for form
function getFormSheet( formName ) {
	var formSheet;
	var activeSheet = SpreadsheetApp.getActiveSpreadsheet();

	// create sheet if needed
	if ( activeSheet.getSheetByName( formName ) == null ) {
      formSheet = activeSheet.insertSheet();
      formSheet.setName( formName );
      isNewSheet = true;
	}
	return activeSheet.getSheetByName( formName );
}


// magic function where it all happens
function insertToSheet( data ){
	var flat = flattenObject( data );
	var keys = Object.keys( flat );
	var formName = data["form_name"];
	var formSheet = getFormSheet( formName );
	var headers = getHeaders( formSheet, keys );
	var values = getValues( headers, flat );
	setHeaders( formSheet, headers );
	setValues( formSheet, values );
	
    if ( emailNotification ) {
		sendNotification( data, getSeetURL() );
	}
}

function getSeetURL() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  return spreadsheet.getUrl();
}

function sendNotification( data, url ) {
	var subject = "A new Elementor Pro Froms subbmition has been inserted to your sheet";
  var message = "A new subbmition has been recived via " + data['form_name'] + " form and inserted into your Google sheet at: " + url;
	MailApp.sendEmail( emailAddress, subject, message, {
		name: 'Automatic Emailer Script'
	} );
}