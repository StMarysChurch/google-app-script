function updateMMVSRegisteration() {
  const EVENTBRITE_API_KEY = "D74CNALISZE5UJQNX3V5";
  const EVENT_ID = 618319549417;
  const API_ENDPOINT = "https://www.eventbriteapi.com/v3/events";
  const SHEET_ID = "1wasW_5V-LF1QZnwaPh3Dt45LEryKJpjfMV7lARP5AbE";
  const PARISH_SHEET_RANGE = "Registerations!B5:C14"
  const ATTENDEE_SHEET_RANGE = "Attendees!A2:D200"

  const CHURCH_1_SHEET_RANGE = "St. Gregorios Orthodox Church (Professional Court)!A2:D200"
  const CHURCH_2_SHEET_RANGE = "St. Thomas Orthodox Church (Toronto)!A2:D200"
  const CHURCH_3_SHEET_RANGE = "St. Mary's Orthodox Syrian Church (Ajax)!A2:D200"
  const CHURCH_4_SHEET_RANGE = "St. Gregorios Indian Orthodox Church (Lakeview)!A2:D200"
  const CHURCH_5_SHEET_RANGE = "St. Thomas Orthodox Church (Ottawa)!A2:D200"
  const CHURCH_6_SHEET_RANGE = "St. John’s Orthodox Church (Hamilton)!A2:D200"
  const CHURCH_7_SHEET_RANGE = "St. George Indian Orthodox Church (London)!A2:D200"

  var options = {
    "headers": {
      'method': 'GET',
      "Authorization": `Bearer ${EVENTBRITE_API_KEY}`
    }
  };

  /**
   * Updates the values in the specified range
   * @param {string} spreadsheetId spreadsheet's ID
   * @param {string} range range of cells of the spreadsheet
   * @param {string} valueInputOption determines how the input should be interpreted
   * @see
   * https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
   * @param {list<list<string>>} _values list of string values to input
   * @returns {*} spreadsheet with updated values
   */
  function updateSats(spreadsheetId, range, valueInputOption, values) {
    // This code uses the Sheets Advanced Service, but for most use cases
    // the built-in method SpreadsheetApp.getActiveSpreadsheet()
    //     .getRange(range).setValues(values) is more appropriate.
    // const values = [...parishes].map(([name, value]) => ([name, value]));

    try {
      let valueRange = Sheets.newValueRange();
      valueRange.range = range;
      valueRange.values = values;

      let batchUpdateRequest = Sheets.newBatchUpdateValuesRequest();
      batchUpdateRequest.data = valueRange;
      batchUpdateRequest.valueInputOption = valueInputOption;

      // var ss = SpreadsheetApp.getActiveSpreadsheet();
      // var sheet = ss.getSheets()[0];
      // var clearRange = sheet.getRange("B5:C14");
      // clearRange.clearContent();

      const result = Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest,
        spreadsheetId);
      return result;
    } catch (err) {
      // TODO (developer) - Handle exception
      console.log('Failed with error %s', err.message);
    }
  };


  try {
    var page = 1;
    var parishes = new Map();
    var totalAttendees = 0;
    var attendee = [];
    var pagination = true;
    while (page == 1 || pagination) {
      console.log(page);
      var attendeesResponse = UrlFetchApp.fetch(`${API_ENDPOINT}/${EVENT_ID}/attendees?status=attending&page=${page}`, options);
      const responseCode = attendeesResponse.getResponseCode();
      if (responseCode !== 200) {
        return '';
      }
      // console.log(attendeesResponse.getResponseCode());
      var attendees = JSON.parse(attendeesResponse.getContentText());
      totalAttendees += Object.keys(attendees.attendees).length;
      pagination = attendees.pagination.has_more_items;
      // console.log(attendees.pagination.has_more_items);
      attendees.attendees.forEach(element => {
        var date = new Date(element.created);
        attendee.push([
          `${date.getFullYear()}/${date.getMonth()}/${date.getDate()}`, element.profile.name, element.answers[0].answer, element.profile.email
        ]);
        if (parishes.has(element.answers[0].answer)) {
          parishes.set(element.answers[0].answer, parishes.get(element.answers[0].answer) + 1);
        } else {
          if (element.answers[0].answer == undefined) {
            console.log(element);
          }
          parishes.set(element.answers[0].answer, 1);
        }
      });
      // console.log(parishes);
      // console.log(attendees);
      // console.log(attendee);
      page++;
    }
    const parishesSorted = [...parishes].sort((a, b) => b[1] - a[1]);
    const church1 = [...attendee].filter(e => e[2]=="St. Gregorios Orthodox Church (Professional Court)");
    const church2 = [...attendee].filter(e => e[2]=="St. Thomas Orthodox Church (Toronto)");
    const church3 = [...attendee].filter(e => e[2]=="St. Mary's Orthodox Syrian Church (Ajax)");
    const church4 = [...attendee].filter(e => e[2]=="St. Gregorios Indian Orthodox Church (Lakeview)");
    const church5 = [...attendee].filter(e => e[2]=="St. Thomas Orthodox Church (Ottawa)");
    const church6 = [...attendee].filter(e => e[2]=="St. John’s Orthodox Church (Hamilton)");
    const church7 = [...attendee].filter(e => e[2]=="St. George Indian Orthodox Church (London)");

    // console.log(church1);
    updateSats(SHEET_ID, PARISH_SHEET_RANGE, 'USER_ENTERED', parishesSorted);
    updateSats(SHEET_ID, ATTENDEE_SHEET_RANGE, 'USER_ENTERED', attendee);
    updateSats(SHEET_ID, CHURCH_1_SHEET_RANGE, 'USER_ENTERED', church1);
    updateSats(SHEET_ID, CHURCH_2_SHEET_RANGE, 'USER_ENTERED', church2);
    updateSats(SHEET_ID, CHURCH_3_SHEET_RANGE, 'USER_ENTERED', church3);
    updateSats(SHEET_ID, CHURCH_4_SHEET_RANGE, 'USER_ENTERED', church4);
    updateSats(SHEET_ID, CHURCH_5_SHEET_RANGE, 'USER_ENTERED', church5);
    updateSats(SHEET_ID, CHURCH_6_SHEET_RANGE, 'USER_ENTERED', church6);
    updateSats(SHEET_ID, CHURCH_7_SHEET_RANGE, 'USER_ENTERED', church7);
  } catch (e) {
    console.log(e);
  }
}
