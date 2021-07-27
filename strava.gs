// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu("Strava App")
    .addItem("Get data", "getStravaActivityData")
    .addToUi();
}

// Get athlete activity data
function getStravaActivityData() {
  // get the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("StravaData");

  var data = callStravaAPI(getLastActivityDate(sheet));

  var stravaData = [];
  data.forEach(function (activity) {
    var activityStartDate = new Date(activity.start_date_local);
    var arr = [];
    arr.push(
      Utilities.formatString(
        "=HYPERLINK(\"https://www.strava.com/activities/%s\", \"%s\")",
        activity.id,
        activity.id
      ),
      activityStartDate,
      activityStartDate.getDay(), // sunday - saturday: 0 - 6
      parseInt(
        Utilities.formatDate(
          activityStartDate,
          SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
          "w"
        )
      ), // week number
      activityStartDate.getMonth() + 1,
      activityStartDate.getFullYear(),
      activity.name,
      activity.type,
      activity.distance / 1000,
      activity.moving_time / 86400,
      activity.elapsed_time / 86400,
      activity.total_elevation_gain,
      (activity.average_speed * 3600) / 1000,
      (activity.max_speed * 3600) / 1000,
      activity.average_cadence,
      activity.average_watts,
      activity.kilojoules,
      activity.average_heartrate,
      activity.max_heartrate,
      activity.max_watts,
      activity.suffer_score
    );
    stravaData.push(arr);
  });

  // paste the values into the Sheet
  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      stravaData.length,
      stravaData[0].length
    )
    .setValues(stravaData);
}

function getLastActivityDate(sheet) {
  if (sheet.getLastRow() > 1) {
    var dataRange = sheet.getDataRange().getValues();
    return new Date(dataRange[sheet.getLastRow() - 1][1]).getTime() / 1000;
  }
  return "1434300800";
}

// call the Strava API
function callStravaAPI(after) {
  // set up the service
  var service = getStravaService();

  if (service.hasAccess()) {
    Logger.log("App has access.");

    var endpoint = "https://www.strava.com/api/v3/athlete/activities";
    var params = Utilities.formatString("?after=%s&per_page=200", after);

    var headers = {
      Authorization: "Bearer " + service.getAccessToken(),
    };

    var options = {
      headers: headers,
      method: "GET",
      muteHttpExceptions: true,
    };

    var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
    return response;
  } else {
    Logger.log("App has no access yet.");

    // open this url to gain authorization from github
    var authorizationUrl = service.getAuthorizationUrl();

    Logger.log(
      "Open the following URL and re-run the script: %s",
      authorizationUrl
    );
  }
}
