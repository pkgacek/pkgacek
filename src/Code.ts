/**
 * @name onOpen
 * @description
 * @returns
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Resumer")
    .addItem("Create resume", "getResume")
    .addItem("Recalculate", "recalculate")
    .addToUi();
}
/**
 * @name recalculate
 * @description This function changes the value of checkbox
 * @returns
 */
function recalculate() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(
    "Settings - Recalculate",
  );
  const cell = sheet.getRange("A2");
  const value = cell.getValue();
  cell.setValue(!value);
}
/**
 * @name getResume
 * @description This function runs createResume function and display the result url
 * @returns
 */
function getResume() {
  const resumeUrl = createResume();
  const ui = SpreadsheetApp.getUi();
  const html = [
    `<a href="${resumeUrl}" target="_blank" rel="noopener noreferrer">Go to your resume</a>`,
  ];
  const htmlOutput = HtmlService.createHtmlOutput(html.join("\n"))
    .setWidth(400)
    .setHeight(250);
  ui.showModalDialog(htmlOutput, "Resume created!");
}
/**
 * @name addResume
 * @param resume
 * @description This function changes the value of resume cell
 * @returns
 */
function addResume(resume: string) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Resume");
  const cell = sheet.getRange("A2");
  cell.setValue(resume);
}
/**
 * @name GETCOLUMNDATA
 * @description Parses all sheets by provided header and returns all data.
 * @param header
 * @param key
 * @param minCount
 * @param maxSize
 * @returns
 * @customfunction
 */
function GETCOLUMNDATA(
  header = "technologies",
  key = true,
  minCount = 1,
  maxSize = 100,
) {
  if (!header) throw new Error("Header is not defined.");
  if (typeof header !== "string") throw new Error("Header must be a string.");
  if (typeof minCount !== "number")
    throw new Error("MinCount must be a number.");
  if (minCount <= 0) throw new Error("MinCount must be greater than 0.");
  if (typeof maxSize !== "number") throw new Error("MaxSize must be a number.");
  if (maxSize <= 0) throw new Error("MaxSize must be greater than 0.");

  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const data = [];
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();
    if (isDataSheet(sheetName)) {
      let values;
      try {
        values = sheet
          .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
          .getValues();
      } catch (e) {
        Logger.log({ e });
      }
      if (values) {
        const headers = values.shift();
        const columnIndex = headers.indexOf(header);
        const enableIndex = headers.indexOf("enable");
        if (columnIndex > 0) {
          const columnData = values
            .filter((row) => row[enableIndex])
            .map((row) => row[columnIndex].split(", "))
            .flat();
          data.push(columnData);
        }
      }
    }
  }

  const formattedData = data.flat().filter((e) => e);

  if (formattedData.length > 0) {
    const countData = filterCount(formattedData, minCount);
    const maxSizedData = countData.splice(0, maxSize);
    const sortedData = maxSizedData.sort((a, b) =>
      a.value.localeCompare(b.value),
    );
    const parsedData = formatArray(sortedData);
    return parsedData;
  }
  return "";
}
/**
 * @name GETTIME
 * @description Parses two dates and gives the total time spent.
 * @param startDate
 * @param endDate
 * @param key
 * @returns
 * @customfunction
 */
function GETTIME(startDate: Date, endDate: Date, key = true) {
  if (!startDate && !endDate) return "";
  if (!(startDate instanceof Date))
    throw new Error("StartDate must be a valid date");
  if (endDate && !(endDate instanceof Date || endDate === null))
    throw new Error("EndDate must be a valid date");
  if (endDate && startDate.getFullYear() > endDate.getFullYear())
    throw new Error("Start year must be less or equal to end year");
  if (
    endDate &&
    (startDate === endDate || startDate.getMonth() === endDate.getMonth())
  )
    return "1 month";

  const df = startDate;
  const dt = endDate ? endDate : new Date();

  let allYears = dt.getFullYear() - df.getFullYear();
  let partialMonths = dt.getMonth() - df.getMonth() + 1;

  if (partialMonths < 0) {
    allYears--;
    partialMonths = partialMonths + 12;
  }
  if (partialMonths === 12) {
    partialMonths = 0;
    allYears++;
  }
  const yearsText = ["year", "years"];
  const monthsText = ["month", "months"];

  const result =
    (allYears == 1
      ? allYears + " " + yearsText[0]
      : allYears > 1
      ? allYears + " " + yearsText[1]
      : "") +
    (allYears && partialMonths ? " and " : "") +
    (partialMonths == 1
      ? partialMonths + " " + monthsText[0]
      : partialMonths > 1
      ? partialMonths + " " + monthsText[1]
      : "");

  return result;
}
