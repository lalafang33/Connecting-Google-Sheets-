const {google} = require('googleapis');
const fs = require('fs');
const credentials = JSON.parse(fs.readFileSync('credentials.json', 'utf8'));

const auth = new google.auth.JWT(
  credentials.client_email,
  credentials.private_key,
  ['https://www.googleapis.com/auth/spreadsheets'],
  );
  const client = new google.auth.JWT(
    credentials.client_email,
    null,
    credentials.private_key,
    ['https://www.googleapis.com/auth/spreadsheets'],
    null
);
const sheets = google.sheets({version: 'v4', auth: client});

// speadsheetID is the ID for the google sheets
const spreadsheetId = '1Bgl3XMqUaCzZ2uqKZSGuIGT9CxJWffvE0uuJKgVoPLw';
const range = 'Sheet1!L3';

client.authorize((error, tokens) => {
  if (error) {
    console.log(error);
    return;
  }

// Get the value in cell L3
const range = 'Sheet1!L3';

sheets.spreadsheets.values.get({
   spreadsheetId, 
   range 
}).then((response) => {
  const x = response.data.values[0][0];
  const numberX = parseInt(x);
  // Get the values in rows 2 to 6 and columns 2 to x (convert to numbers)  
  // const valuesRange = `Sheet1!B2:H6`;
  const valuesRange = `Sheet1!B2:${String.fromCharCode(64 + (numberX + 1))}6`;
  sheets.spreadsheets.values.get({ 
    spreadsheetId, 
    range: valuesRange, 
  }).then((response) => {
    const values = response.data.values.map((row) => row.map((cell) => Number(cell)));
    // Add the values
    const sum = values.reduce((total, row) => total + row.reduce((rowTotal, cell) => rowTotal + cell), 0);
    // Write the result to cell L4
    const resultRange = 'Sheet1!L4';
    const result = [[sum]];
    sheets.spreadsheets.values.update({ spreadsheetId, range: resultRange, valueInputOption: 'RAW', resource: { values: result }
  });
  })
})
});
