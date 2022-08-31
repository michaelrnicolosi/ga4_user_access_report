function listUserInformation() {

  // Read the inputs from the sheet
  var properties = SpreadsheetApp.getActive().getSheetByName('user_access_report').getRange(3, 2, 100, 1).getDisplayValues();
  var dimensionSelected = SpreadsheetApp.getActive().getSheetByName('user_access_report').getRange(3, 4, 11, 1).getDisplayValues();
  var startDate = SpreadsheetApp.getActive().getSheetByName('user_access_report').getRange(3, 8, 1, 1).getDisplayValues();
  var endDate = SpreadsheetApp.getActive().getSheetByName('user_access_report').getRange(4, 8, 1, 1).getDisplayValues();

  // Dimension lookup table
  var dimensionLookup = {
    0: "accessMechanism",
    1: "accessedPropertyId",
    2: "accessedPropertyName",
    3: "costDataReturned",
    4: "epochTimeMicros",
    5: "mostRecentAccessEpochTimeMicros",
    6: "propertyUserLink",
    7: "reportType",
    8: "revenueDataReturned",
    9: "userEmail",
    10: "userIP"
  }

  // Store which dimensions are selected
  var dimensions = [];

  for (let i = 0; i < dimensionSelected.length; i++) {
    if (dimensionSelected[i] == "TRUE") {
      dimensions.push({
        dimension_name: dimensionLookup[i]
      });
    }
  }


  var allDimensions = [];
  var allMetrics = [];

  for (let i = 0; i < properties.length; i++) {
    if (properties[i] != '') {
      var propertyIdPath = 'properties/' + properties[i];


      // Make requests to the GA4 Admin API
      var accessInfo = AnalyticsAdmin.Properties.runAccessReport(resource = {
        date_ranges: {
          start_date: startDate[0][0],
          end_date: endDate[0][0]
        },
        dimensions: dimensions,
        metrics: [{
          metric_name: 'accessCount'
        }]
      },
        parent = propertyIdPath
      );

      for (let j = 0; j < accessInfo.rows.length; j++) {
        allDimensions.push(accessInfo.rows[j].dimensionValues);
        allMetrics.push(accessInfo.rows[j].metricValues);
      }
    }
  }

  var lastRow = allDimensions.length - 1000;
  var results = SpreadsheetApp.getActiveSpreadsheet();
  if (SpreadsheetApp.getActive().getSheetByName('results') != null) {
    results.deleteSheet(results.getSheetByName('results'));
  }
  results.insertSheet('results');

  if (allDimensions.length > 1000) {
    results.insertRowsAfter(1000, lastRow);
  }

  var setResults = SpreadsheetApp.getActive().getSheetByName('results');
  for (let i = 0; i < dimensions.length; i++) {
    var dimensionColumnIndex = i + 1;
    var metricColumnIndex = dimensions.length + 1;
    setResults.getRange(1, dimensionColumnIndex, 1, 1).setValue(dimensions[i].dimension_name);
    setResults.getRange(1, metricColumnIndex, 1, 1).setValue("accessCount");

    for (var j = 0; j < allDimensions.length; j++) {
      let dimensionValue = allDimensions[j][i].value;
      let metricValue = allMetrics[j][0].value;

      if (dimensions[i].dimension_name == "mostRecentAccessEpochTimeMicros" || dimensions[i].dimension_name == "epochTimeMicros") {
        dimensionValue = new Date(Number(dimensionValue) / 1000).toUTCString();

      }
      var startingRow = j + 2;
      setResults.getRange(startingRow, dimensionColumnIndex, 1, 1).setValue(dimensionValue);
      setResults.getRange(startingRow, metricColumnIndex, 1, 1).setValue(metricValue);
    }

  }
}
