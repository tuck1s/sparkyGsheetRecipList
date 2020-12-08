function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Recipient List')
    .addItem('Set API Key', 'setAPIKey')
    .addItem('View Recipient Lists', 'viewRecipientLists')
    .addSeparator()
    .addItem('Upload Recipient List', 'uploadRecipientList')
    .addToUi();
  defaultURL(); // ensure we begin with a default SparkPost service URL
}

// SparkPost API paths
const recipListPath = '/api/v1/recipient-lists';

// Items  stored in user-specific Google Scripts cache service
const configApiKey = 'apiKey';

function getConfig(c) {
  var cache = CacheService.getUserCache();
  return (cache.get(c));
}

function setConfig(c, val) {
  var cache = CacheService.getUserCache();
  cache.put(c, val)
}

function setConfigWithExpiry(c, val, expiry) {
  var cache = CacheService.getUserCache();
  cache.put(c, val, expiry)
}

// API Key management
function setAPIKey() {
  const ui = SpreadsheetApp.getUi();
  const expiry = 3600;
  var response = ui.prompt('Set SparkPost API key', 'Should have Recipients List: Read/Write permission.\nFor security reasons this will be held for no more than ' + expiry.toString() + ' seconds (and may be purged sooner).\n', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    k = response.getResponseText()
    setConfigWithExpiry(configApiKey, k, expiry);
  }
}

function validKey(k) {
  return (k !== null && k.length >= 4) // Simple validity check
}

function redactedKey(k) {
  if (validKey(k)) {
    // Redact most of the key for display, like SparkPost UI does
    var tail = '*'.repeat(k.length - 4);
    return (k.slice(0, 4) + tail);
  }
  return ('Invalid');
}

// SparkPost URL management
function getSparkPostURL() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var endpointCell = sheet.getRange('B4');
  switch(endpointCell.getValue()) {
    case 'US':
      return ('https://api.sparkpost.com');
    case 'EU':
      return ('https://api.eu.sparkpost.com');
    default:
      endpointCell.setValue('US');
      return ('https://api.sparkpost.com');
  }
}

// Fetch your current recipient lists from SparkPost and display them in a table
function viewRecipientLists() {
  const ui = SpreadsheetApp.getUi(); // Or DocumentApp or FormApp.
  var apiKey = getConfig(configApiKey);
  if (!validKey(apiKey)) {
    ui.alert('API key not set');
    return;
  }
  var options = {
    headers: {
      'Content-Type': 'application/json',
      'Authorization': apiKey,
    },
    muteHttpExceptions: true,
  };
  var fullPath = getSparkPostURL() + recipListPath;
  var response = UrlFetchApp.fetch(fullPath, options);
  if (response.getResponseCode() !== 200) {
    ui.alert('API key: ' + redactedKey(apiKey), response.getContentText(), ui.ButtonSet.OK);
    return;
  }
  var res = JSON.parse(response.getContentText())['results'];
  var display = htmlTableFromJSON(res, ['id', 'name', 'total_accepted_recipients']);
  var htmlOutput = HtmlService
    .createHtmlOutput(display)
    .setWidth(800)
    .setHeight(800);
  ui.showModelessDialog(htmlOutput, redactedKey(apiKey));
}

function htmlTableFromJSON(res, tableHeadings) {
  // report the number of results as the title of the table
  const lists = `Results: ${res.length.toString()}`;
  var tableRows = [];
  for (const l of res) {
    var tr = [];
    // Extract from dictionary keyed by tableHeadings into an array
    for (var i = 0; i < tableHeadings.length; i++) {
      tr.push(l[tableHeadings[i]]);
    }
    tableRows.push(tr);
  }
  const head = shadedTableStyle('arial, sans-serif', '#dddddd');
  const table = htmlTable(lists, tableHeadings, tableRows);
  return (html(head, table));
}

// Build HTML using templates
function html(headContent, bodyContent) {
  return (`
<html>
<head>
${headContent}
</head>
<body>
${bodyContent}
</body>
</html>`);
}

function shadedTableStyle(font, shadedColor) {
  return (`
<style>
table {
  font-family: ${font};
  border-collapse: collapse;
  width: 100%;
}
td, th {
  border: 1px solid ${shadedColor};
  text-align: left;
  padding: 8px;
}
tr:nth-child(even) {
  background-color: ${shadedColor};
}
</style>`);
}

function htmlTable(title, heading, data) {
  var thdr = `<tr>`;
  for (const i of heading) {
    thdr += `<th>${i}</th>`;
  }
  thdr += `</tr>`;

  var tbody = '';
  for (const i of data) {
    tbody += `<tr>`;
    for (const j of i) {
      tbody += `<td>${j}</td>`;
    }
    tbody += `</tr>`;
  }
  return (`
<h2>${title}</h2>
<table>
${thdr}
${tbody}
</table>`);
}

// Upload the current sheet to SparkPost recipient list
function uploadRecipientList() {
  const ui = SpreadsheetApp.getUi();
  var apiKey = getConfig(configApiKey);
  if (validKey(apiKey)) {
    // Make a POST request with a JSON payload.
    const data = getData();
    const options = {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': apiKey,
      },
      muteHttpExceptions: true,
      payload: JSON.stringify(data)
    };
    const fullPath = getSparkPostURL() + recipListPath;
    const response = UrlFetchApp.fetch(fullPath, options);
    if (response.getResponseCode() === 200) {
      r = JSON.parse(response.getContentText())['results'];
      r_acc = r['total_accepted_recipients'];
      r_rej = r['total_rejected_recipients'];
      ui.alert(r['id'] + ' : "' + r['name'] + '"',
        'Total accepted recipients: ' + r_acc.toString() + '\n\nTotal rejected recipients: ' + r_rej.toString(),
        ui.ButtonSet.OK);
    } else {
      // just report any errors back to the user in raw form
      ui.alert(response.getContentText());
    }
  } else {
    ui.alert('Please set SparkPost API key');
  }
}

// Read values from specific cell locations
function getData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  if (data[0][0] != 'list_id') {
    return ('list_id field not in the expected place', null, null);
  }
  const list_id = data[0][1];

  if (data[1][0] != 'list_name') {
    return ('list_name field not in the expected place', null, null);
  }
  const list_name = data[1][1];

  if (data[2][0] != 'list_description') {
    return ('list_description field not in the expected place', null, null);
  }
  const list_description = data[2][1];

  const list_hdr = data[4];
  // Check the mandatory header column names are as expected
  if (list_hdr[0] != 'email.address' || list_hdr[1] != 'email.name') {
    return ('email.address and email.name header columns not in the expected place', null, null);
  }
  // Collect the optional substitution_data column names
  const subDataName = list_hdr.slice(2);
  // Form a list of recipients
  var recipients = [];
  for (var row = 5; row < data.length; row++) {
    Logger.log(data[row]);
    subData = data[row].slice(2);
    if (subData.length != subDataName.length) {
      return ('Row ${i+1} data length does not match substitution_data header', null, null);
    }
    // Build a SparkPost recipient object, adding sub data if present
    var recip = {
      'address': {
        'email': data[row][0],
        'name': data[row][1],
      }
    };
    if (subDataName.length) {
      var sub = {};
      subDataName.forEach((key, i) => sub[key] = subData[i]);
      recip['substitution_data'] = sub;
    }
    recipients.push(recip);
  }
  const r = {
    'id': list_id,
    'name': list_name,
    'description': list_description,
    'recipients': recipients
  };
  return (r);
}

