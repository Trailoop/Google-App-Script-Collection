
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


function getCounty(cell) {
  Utilities.sleep(9000);
  var response = Maps.newGeocoder().geocode(cell);
  var addressComponents =  response.results[0].address_components;
  var county = ""
  for (var i = 0; i < addressComponents.length; i++) {
      var types = addressComponents[i].types.join(",")
      if (types == "administrative_area_level_2,political") {
      
        county = addressComponents[i].short_name;
      }
  }
  
  return county;
  
}

function test(cell){
Utilities.sleep(9000)
var response = Maps.newGeocoder().geocode(cell);
var lat = response.results[0].geometry.location.lat;
var lng = response.results[0].geometry.location.lng;
var result = lat + ":" + lng;
return result.toString();

}

function getAddress(cell){
Utilities.sleep(9000)
var response = Maps.newGeocoder().geocode(cell);
var address = response.results[0].formatted_address;
return address;

}

function getAllCounties() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var range = sheet.getRange(sheet.getDataRange().getLastRow() - sheet.getDataRange().getNumRows(), sheet.getDataRange().getLastColumn(), sheet.getDataRange().getNumRows(), 1); 
 var locationInfo = range.getValues();
 var response,county;
 var values = [];
   for (var i = 0; i < locationInfo.length; i++) {
    response = Maps.newGeocoder().geocode(locationInfo[i]);
 
 var addressComponents = response.results[0].address_components;

  for (var i = 0; i < addressComponents.length; i++) {
      var types = addressComponents[i].types.join(",")
      if (types == "administrative_area_level_2,political") {
        county = addressComponents[i].long_name
        values.push(county);
      }
   }
   Utilities.sleep(5000);
  }
  
   return values;
   
}

