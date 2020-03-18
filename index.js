const fs = require('fs');
const path = require('path');
const readline = require('readline');
const {google} = require('googleapis');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.resolve('./state/token.json');

function addZero(n) {
  return n < 10 ? "0" + n : "" + n;
}

function formatDate(date) {
  var day = date.getDate();
  var monthIndex = date.getMonth();
  var year = date.getFullYear();

  return year + '-' + addZero(monthIndex +1) + '-' + addZero(day);
}

// Action 1: Run the downloadscript, marking the hours as exported in each sheet
// to export hours again, mark the weeks as editable
//   requires you to specify the end week
// calculate the date range for the year/week number
// Loop over all sprint sheets sorted by sprint number 
// If the weeks date < endweek date && !readonly(week)
  // generate csv rows for the week
  // mark week as readonly 



// Action 2: given a timetell export, verify it against the data in the excel 


async function getEmployeeInfo(sheets) {
  const data = await sheets.spreadsheets.values.get({
    spreadsheetId: "1ZCp_DD6_c25bjXdoGEYOpQlvfFXb0ARTnh-EtzJr0e8",
    range: "'medewerkers'!A2:D" 
  });

  let hasError = false
  const retVal = data.data.values.reduce((p, c) => {
    const voornaam = c[0].toLowerCase();
    if (p[voornaam]) {
      console.error(`Employee ${voornaam} defined more than once`)
    }
    p[voornaam] = c[3];
    return p
  }, {});
  return [hasError, retVal]
}

async function getProjectInfo(sheets) {
  const data = await sheets.spreadsheets.values.get({
    spreadsheetId: "1ZCp_DD6_c25bjXdoGEYOpQlvfFXb0ARTnh-EtzJr0e8",
    range: "'proj_en_act'!A2:D" 
  })
  
  let hasError = false
  const retVal = data.data.values.reduce((p, c) => {
    if (p[c[3]]) {
      hasError = true
      console.error(`Employee ${c[3]} defined more than once`)
    }
    p[c[3]] = {
      id: c[1], 
      is_act: c[2] === "1"
    }; 
    return p
  }, {});
  return [hasError, retVal]
}

async function getSprintTabsInDescendingOrder(sheets, spreadsheetId) {
  const sheetInfo = await sheets.spreadsheets.get({ spreadsheetId });

  return sheetInfo.data.sheets
    .map(s => ({title: s.properties.title, sheetId: s.properties.sheetId, protectedRanges: s.protectedRanges}))
    .filter(s => /^Sprint [0-9]+$/.test(s.title))
    .sort((a, b) => (+b.title.split(" ")[1]) - (a.title.split(" ")[1]));
}

const reportedErrors = {};
function reportError(msg) {
  if (reportedErrors[msg] === undefined) {
    console.error(msg)
    reportedErrors[msg] = true;
  }
}

function isReadOnly(protectedRanges, startCol, nextStartCol) {
  if (protectedRanges === undefined) {
    return false;
  }
  const result = protectedRanges
    .some(({range}) => 
      range.startColumnIndex <= nextStartCol && 
      range.endColumnIndex > startCol && 
      (range.endRowIndex === undefined || range.endRowIndex >= (HOUR_DATA_START_ROW - 1))
    )
//   for (const {range} of protectedRanges) {
//     console.error(`  r.startColumnIndex (${range.startColumnIndex}) <= nextStartCol (${nextStartCol}) && 
//   r.endColumnIndex (${range.endColumnIndex}) > startCol (${startCol}) &&
//   (range.endRowIndex === undefined || r.endRowIndex (${range.endRowIndex}) >= HOUR_DATA_START_ROW (${HOUR_DATA_START_ROW})) === ${range.startColumnIndex <= nextStartCol && 
//     range.endColumnIndex > startCol &&
//     (range.endRowIndex >= HOUR_DATA_START_ROW || range.endRowIndex === undefined)}
// `)
//   }
  return result;
}

function getWeekNumber(srcDate) {
  return getWeekYear(srcDate) * 100 + getWeek(srcDate);
}

function getWeekYear(srcDate) {
  var date = new Date(srcDate.getTime());
  date.setDate(date.getDate() + 3 - (date.getDay() + 6) % 7);
  return date.getFullYear();
}

function getWeek(srcDate) {
  var date = new Date(srcDate.getTime());
  date.setHours(0, 0, 0, 0);
  // Thursday in current week decides the year.
  date.setDate(date.getDate() + 3 - (date.getDay() + 6) % 7);
  // January 4 is always in week 1.
  var week1 = new Date(date.getFullYear(), 0, 4);
  // Adjust to Thursday in week 1 and count number of weeks from date to week1.
  return 1 + Math.round(((date.getTime() - week1.getTime()) / 86400000
                        - 3 + (week1.getDay() + 6) % 7) / 7);
}

async function run(spreadsheetIds, endweek) {
  const aggregateData = {}
  for (const spreadsheetId of spreadsheetIds) {
    await downloadHours({aggregateData, spreadsheetId, endweek})
  }
  printData(aggregateData);
}


const HOUR_DATA_HEADER_ROW = 52
const HOUR_DATA_START_ROW = 53
const WEEK_START_COL = [6, 7, 12, 16]

async function downloadHours({aggregateData, spreadsheetId, endweek}) {
  const auth = await authorize(JSON.parse(fs.readFileSync('credentials.json', "utf-8")));
  const sheets = google.sheets({version: 'v4', auth});
  const [empError, employees] = await getEmployeeInfo(sheets);
  const [projeErro, projs] = await getProjectInfo(sheets);
  if (empError || projeErro) {
    return
  }
  const tabs = await getSprintTabsInDescendingOrder(sheets, spreadsheetId);
  for (const tab of tabs) {
    const valueSegment = `'${tab.title}'!A${HOUR_DATA_START_ROW}:S`;
    const headerSegment = `'${tab.title}'!A${HOUR_DATA_HEADER_ROW}:S${HOUR_DATA_HEADER_ROW}`;

    const header = (await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: headerSegment,
      valueRenderOption: "UNFORMATTED_VALUE",
      dateTimeRenderOption: "SERIAL_NUMBER"
    })).data.values[0];

    const weeks = WEEK_START_COL
      .slice(0, -1)
      .map((c, i) => [
        getWeekNumber(parseSheetDate(header[c])),
        WEEK_START_COL[i],
        WEEK_START_COL[i + 1]
      ])
      .filter(([weeknum]) => weeknum <= endweek)
      .filter(([weeknum, weekStart, weekEnd]) => !isReadOnly(tab.protectedRanges, weekStart, weekEnd))
    

    console.error(`Exporting ${tab.title} ${JSON.stringify(weeks.map(x => x[0]))}`)

    if (weeks.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: weeks.map(([weeknum, weekStart, weekEnd]) => ({
            "addProtectedRange": {
              "protectedRange": {
                description: `The cells for week ${weeknum} were already exported to timetell`,
                range: {
                  sheetId: tab.sheetId,
                  startColumnIndex: weekStart,
                  endColumnIndex: weekEnd,
                  startRowIndex: HOUR_DATA_START_ROW  - 1
                },
                warningOnly: true
              }
            }
          }))
        }
      });
    
      if (header[0] !== 'Project\n(voor werk dat \ntussendoor komt)' || header[18] !== 'TOTAAL') {
        console.error(`values segment (${valueSegment}) mismatch. It does not start with 'Project\\n(voor werk dat \\ntussendoor komt)' or it does not end with 'TOTAAL'.`)
        console.error(header.map((v,i) => "  " + i +": " +JSON.stringify(v)).join("\n"))
      } else {
        const values = (await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: valueSegment,
          valueRenderOption: "UNFORMATTED_VALUE",
          dateTimeRenderOption: "SERIAL_NUMBER"
        })).data.values;
        for (const row of values) {
          if (row === undefined) {
            continue
          }
          const proj = row[17].split(" - ")[0].trim();
          const person = employees[row[3].trim().toLowerCase()];
          const remark = row[2].trim();
          const projInfo = projs[proj];
          if (projInfo === undefined) {
            if (row[17] !== "NO EPIC!") {
              reportError(`project ${JSON.stringify(row[17])} is not found! lookup=${JSON.stringify(proj)}, keys: ${Object.keys(projs)}`)
            }
            continue;
          }
          if (row[3].trim() === "") {
            continue
          }
          if (person === undefined) {
            reportError(`employee ${JSON.stringify(row[3])} is not found!`)
            continue
          }

          for (const [weeknum, weekStart, weekEnd] of weeks) {
            for (let i = weekStart; i<weekEnd; i++) {
              if (row[i] === "") {
                continue
              }
              const key = JSON.stringify([
                person, 
                projInfo.id, 
                projInfo.is_act, 
                formatDate(parseSheetDate(header[i]))
              ]);
              if (aggregateData[key] === undefined) {
                aggregateData[key] = [0, []]
              }
              aggregateData[key][0] += row[i];
              aggregateData[key][1].push(remark);
            }
          }
        }
      }
    }
  }

}

function printData(aggregateData) {
  console.log(["MED_NR","ACT_NR","PRJ_NR","DATUM","UREN"].join(","))
  for (const key in aggregateData) {
    const [medewerker, code, is_act, datum] = JSON.parse(key)
    const [uren, opmerking] = aggregateData[key]
    console.log([
      medewerker,
      is_act ? code : "",
      is_act ? "" : code,
      datum,
      uren,
      opmerking.join(", ")
    ].map(x => JSON.stringify(x)).join(";"))
  }
}

function parseSheetDate(num) {
  return new Date(1900, 0, num - 1);
}

run(process.argv.slice(3), +process.argv[2]).catch(e => console.error(e))  

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
async function authorize(credentials) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  try {
    const token = fs.readFileSync(TOKEN_PATH, "utf-8");
    oAuth2Client.setCredentials(JSON.parse(token));
    return oAuth2Client;
  } catch (e) {
    console.log("tokenerror", e)
    return await getNewToken(oAuth2Client);
  }
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.error('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  return new Promise((resolve, reject) => {
    rl.question('Enter the code from that page here: ', (code) => {
      rl.close();
      oAuth2Client.getToken(code, (err, token) => {
        if (err) {
          reject(err)
        } else {
          oAuth2Client.setCredentials(token);
          // Store the token to disk for later program executions
          fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
            if (err) {
              reject(err)
            } else {
              console.error('Token stored to', TOKEN_PATH);
            }
          });
          resolve(oAuth2Client);  
        }
      });
    });  
  })
}
