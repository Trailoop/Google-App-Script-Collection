
function geocode_Addresses() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var row = sheet.getActiveCell().getRow();
 var column = sheet.getActiveCell().getColumn()-1;
 var numRows = sheet.getDataRange().getLastRow() - row;
 var range = sheet.getRange(row, column, numRows, 1);
 var locationInfo = range.getValues();
 var geoResults, lat, lng;
 var destinationRange = sheet.getRange(row, column+2, numRows, 1);
 var values = [];
   for (var i = 0; i < locationInfo.length; i++) {
    geoResults = Maps.newGeocoder().geocode(locationInfo[i]);

    // Get the latitude and longitude
    lat = geoResults.results[0].geometry.location.lat;
    lng = geoResults.results[0].geometry.location.lng;
    var result = lat + ":" + lng;
    values.push(result.toString());
    Utilities.sleep(5000); 
    }
    
    return values;


}

function getAddressFormatted() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var row = sheet.getActiveCell().getRow();
 var column = sheet.getActiveCell().getColumn()-1;
 var numRows = sheet.getDataRange().getLastRow() - row;
 var range = sheet.getRange(row, column, numRows, 1);
 var locationInfo = range.getValues();
 var geoResults, address;
 var destinationRange = sheet.getRange(row, column+2, numRows, 1);
 var values = [];
   for (var i = 0; i < locationInfo.length; i++) {
    geoResults = Maps.newGeocoder().geocode(locationInfo[i]);
 
 geoResults = Maps.newGeocoder().geocode(locationInfo[i]);
 
    // Get the latitude and longitude
    
    address = geoResults.results[0].formatted_address;
    values.push(address);
    Utilities.sleep(1000); 
    }
    
    return values;


}
