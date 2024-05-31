// Schema is
// timestamp, type, url, name, title, company

function signin() {
  chrome.identity.getAuthToken({ interactive: true }, function (token) {
    if (token === undefined) {
      console.log("Error authenticating with Google Drive");
    } else {
      gapi.load("client", function () {
        gapi.client.setToken({ access_token: token });
        gapi.client.load("drive", "v3", function () {
          gapi.client.load("sheets", "v4", function () {
            // main();
          });
        });
      });
    }
  });
}
signin();

const SPREADSHEET_ID = "DkhlalVlc1TLNfdIj2i9mRJni5xp3ac9_0iU_PeJy3Q";

async function getSheetIdFromTitle(title: string): Promise<number | null> {
  try {
    const spreadsheetId = SPREADSHEET_ID; // Replace with your actual spreadsheet ID
    const response = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId,
    });

    const sheets = response.result.sheets;

    if (!sheets) return null;

    for (const sheet of sheets) {
      if (sheet.properties && sheet.properties.title === title) {
        const sheetId = sheet.properties.sheetId;
        console.log("FOUND SHEET ID: ", sheetId);
        return sheetId ?? null;
      }
    }

    console.log("Sheet not found");
  } catch (error) {
    console.error("Error finding sheet:", error);
  }

  return null;
}

async function createSheet(sheetTitle: string): Promise<number | null> {
  try {
    console.log(`createSheet(${sheetTitle})`);
    const spreadsheetId = SPREADSHEET_ID;

    // First, check if the sheet already exists
    const existingSheet = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId,
    });

    const sheets = existingSheet.result.sheets;
    console.log(`createSheet - sheets: ${JSON.stringify(sheets)}`);
    const sheetExists = sheets?.some(
      (sheet) => sheet.properties?.title === sheetTitle
    );

    console.log(`createSheet - sheetExists: ${sheetExists}`);

    if (sheetExists) {
      console.log('Sheet "Personal CRM" already exists.');
      if (!sheets) return null;
      for (const sheet of sheets) {
        if (sheet.properties?.title === sheetTitle) {
          const sheetId = sheet.properties.sheetId;
          console.log(`Sheet ID: ${sheetId}`);
          return sheetId ?? null;
        }
      }
    }

    // If the sheet does not exist, create a new one
    const requests = [
      {
        addSheet: {
          properties: {
            title: sheetTitle,
          },
        },
      },
    ];

    const batchUpdateRequest = {
      requests,
    };

    // Send the request to create the sheet
    const response = await gapi.client.sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: batchUpdateRequest,
    });

    console.log('Sheet "Personal CRM" created successfully.');
    let sheetId: number | null = null;
    const replies = response.result.replies;
    if (replies && replies[0]?.addSheet?.properties?.sheetId) {
      sheetId = replies[0].addSheet.properties.sheetId;
    }
    return sheetId;
  } catch (error) {
    console.error("Error creating sheet:", error);
    return null;
  }
}

function getSheetHeaders(
  sheetId: number,
  sheetTitle: string
): Promise<string[]> {
  return gapi.client.sheets.spreadsheets.values
    .get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetTitle}!1:1`, // Assuming headers are in the first row
    })
    .then((response) => {
      if (!response.result.values) {
        return [];
      } else {
        const values: string[] = response.result.values[0];
        return values; // Returns the headers as an array
      }
    });
}

function setLoading(status: boolean) {
  const loadingElement = document.getElementById("loader");
  if (loadingElement) {
    loadingElement.style.display = status ? "block" : "none";
  }
}

function appendToSheet(row: string[], sheetTitle: string) {
  var rangeEnd = String.fromCharCode("A".charCodeAt(0) + row.length);
  if (sheetTitle) {
    var appendParams = {
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetTitle}!A:${rangeEnd}`,
      valueInputOption: "RAW",
    };
  } else {
    var appendParams = {
      spreadsheetId: SPREADSHEET_ID,
      range: "A:" + rangeEnd,
      valueInputOption: "RAW",
    };
  }

  var valueRangeBody = {
    majorDimension: "ROWS",
    values: [row],
  };

  return gapi.client.sheets.spreadsheets.values.append(
    appendParams,
    valueRangeBody
  );
}

async function getTabContent(): Promise<string> {
  return new Promise((resolve, reject) => {
    chrome.tabs.executeScript(
      {
        code: "document.documentElement.outerHTML;",
      },
      function (results) {
        if (chrome.runtime.lastError) {
          reject(
            new Error(
              "Error in executing script: " + chrome.runtime.lastError.message
            )
          );
        } else if (results && results[0]) {
          resolve(results[0]);
        } else {
          reject(new Error("No result returned from script execution"));
        }
      }
    );
  });
}

interface WeWorkData {
  title: string;
  address: string;
  availableSeats: number;
  datetime: string;
}

async function parseWeWorkData(): Promise<WeWorkData[]> {
  const weWorkSelector = "wework-booking-desk-memberweb .list-unstyled .col-12";
  return new Promise((resolve, reject) => {
    getTabContent()
      .then((html) => {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, "text/html");
        const weWorkElements = doc.querySelectorAll(weWorkSelector);
        const weWorkData: WeWorkData[] = [];

        weWorkElements.forEach((element) => {
          const titleElement = element.querySelector("h3");
          const addressElement = element.querySelector("p");
          const availableSeatsElement = element.querySelector("span");

          if (titleElement && addressElement && availableSeatsElement) {
            const title = titleElement.textContent || "";
            const address = addressElement.textContent || "";
            const availableSeats =
              parseInt(availableSeatsElement.textContent || "", 10) || 0;
            const datetime = new Date().toISOString();

            weWorkData.push({ title, address, availableSeats, datetime });
          }
        });

        resolve(weWorkData);
      })
      .catch((error) => {
        reject(error);
      });
  });
}

async function main() {
  setLoading(true);
  try {
    const sheetTitle = "WeWork Data";
    let sheetId = await getSheetIdFromTitle(sheetTitle);
    if (!sheetId) {
      sheetId = await createSheet(sheetTitle);
    }

    if (sheetId) {
      const weWorkData = await parseWeWorkData();
      for (const data of weWorkData) {
        const row = [
          data.title,
          data.address,
          data.availableSeats.toString(),
          data.datetime,
        ];
        await appendToSheet(row, sheetTitle);
      }
      console.log("Data appended successfully");
    } else {
      console.error("Failed to get or create sheet");
    }
  } catch (error) {
    console.error("Error in main function:", error);
  } finally {
    setLoading(false);
  }
}

document.addEventListener("DOMContentLoaded", function () {
  const logButton = document.getElementById("logCapacityButton");
  if (logButton) {
    logButton.addEventListener("click", main);
  }
});

document.getElementById("openSheet")?.addEventListener("click", function () {
  chrome.tabs.create({
    url: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}`,
  });
});
