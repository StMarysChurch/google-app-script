function updateMMVSRegisteration() {
  // Set EVENTBRITE_API_KEY, EVENT_ID, SHEET_ID and SHEET_RANGE
  const EVENTBRITE_API_KEY = "";
  const EVENT_ID = 0;
  const API_ENDPOINT = "https://www.eventbriteapi.com/v3/events";
  const SHEET_ID = "";
  const SHEET_RANGE = "Sheet1!B5:C14"

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
  function updateSats(spreadsheetId, range, valueInputOption, parishes) {
    // This code uses the Sheets Advanced Service, but for most use cases
    // the built-in method SpreadsheetApp.getActiveSpreadsheet()
    //     .getRange(range).setValues(values) is more appropriate.
    const values = [...parishes].map(([name, value]) => ([name, value]));

    try {
      let valueRange = Sheets.newValueRange();
      valueRange.range = range;
      valueRange.values = values;

      let batchUpdateRequest = Sheets.newBatchUpdateValuesRequest();
      batchUpdateRequest.data = valueRange;
      batchUpdateRequest.valueInputOption = valueInputOption;

      const result = Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest,
        spreadsheetId);
      return result;
    } catch (err) {
      // TODO (developer) - Handle exception
      console.log('Failed with error %s', err.message);
    }
  };

  // Fetches attendes info from eventbrite event
  // Then aggreates the answeres from eventbrite signup form
  try {
    var attendeesResponse = UrlFetchApp.fetch(`${API_ENDPOINT}/${EVENT_ID}/attendees?status=attending`, options);
    const responseCode = attendeesResponse.getResponseCode();
    if (responseCode !== 200) {
      return '';
    }
    console.log(attendeesResponse.getResponseCode());
    var attendees = JSON.parse(attendeesResponse.getContentText());
    console.log(attendees);
    if (!attendees.pagination.has_more_items) {
      const totalAttendees = Object.keys(attendees.attendees).length;
      const parishes = new Map();
      attendees.attendees.forEach(element => {
        if (parishes.has(element.answers[0].answer)) {
          parishes.set(element.answers[0].answer, parishes.get(element.answers[0].answer) + 1)
        } else {
          parishes.set(element.answers[0].answer, 1)
        }
      });
      console.log(parishes);
      updateSats(SHEET_ID, SHEET_RANGE, 'USER_ENTERED', parishes);
    }
  } catch (e) {
    console.log(e);
  }
}