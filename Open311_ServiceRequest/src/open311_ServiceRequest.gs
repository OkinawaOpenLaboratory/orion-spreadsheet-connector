function onEditIssues() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Open311ServiceRequest');
  var lastRow = sheet.getLastRow();
  var idName = "service-request" + lastRow;
  sheet.getRange(lastRow, 1).setValue(idName);
  var id = sheet.getRange(lastRow, 1).getValue();
  var description = sheet.getRange(lastRow, 2).getValue();
  var location = sheet.getRange(lastRow, 3).getValue();
  var splitRow = location.indexOf(',');
  var latitude = location.slice(0, splitRow);
  var longitude = location.slice(splitRow + 1)
  var mediaUrl = sheet.getRange(lastRow, 4).getValue();
  var serviceCode = sheet.getRange(lastRow, 5);
  var serviceName = sheet.getRange(lastRow, 6).getValue();
  var status = 'open';
  var requestedDatetime = sheet.getRange(lastRow, 6);
  var postData = {
    "id": id,
    "type": "Open311ServiceRequest",
    "address": "",
    "agency_responsible": "",
    "alternateName": "",
    "areaServed": "",
    "dataProvider": "",
    "dateCreated": "",
    "dateModified": "",
    "description": description,
    "device_id": "",
    "email": "",
    "expected_datetime": "",
    "first_name": "",
    "jurisdiction_id": "",
    "last_name": "",
    "location": {
      "type": "Point",
      "value": [ Number(latitude), Number(longitude)]
    },
    "media_url": {
      "type": "URL",
      "value": mediaUrl
    },
    "name": "",
    "phone": "",
    "requested_datetime": {
      "type": "DateTime",
      "value": requestedDatetime
    },
    "seeAlso": "",
    "service_code": serviceCode,
    "service_name": serviceName,
    "service_notice": "",
    "service_request_id": "",
    "source": "",
    "status": status,
    "status_notes": "",
    "updated_datetime": {
      "type": "DateTime",
      "value": ""
    }
  }
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(postData),
  };
  UrlFetchApp.fetch("https://<orion-host>:1026/v2/entities?options=keyValues", options);
}
