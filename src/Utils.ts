type Row = Record<string, any>;
const LINK_REGEXP = new RegExp(
  /\[(.+)\]\(((tel:[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,8})|(mailto:[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)|(((?:https?)|(?:ftp)):\/\/(?:www\.|(?!www))[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|www\.[a-zA-Z0-9][a-zA-Z0-9-]+[a-zA-Z0-9]\.[^\s]{2,}|https?:\/\/(?:www\.|(?!www))[a-zA-Z0-9]+\.[^\s]{2,}|www\.[a-zA-Z0-9]+\.[^\s]{2,}))\)/,
);
const CONSENT =
  "I hereby give consent for my personal data included in the application to be processed for the purposes of the recruitment process in accordance with Art. 6 paragraph 1 letter a of the Regulation of the European Parliament and of the Council (EU) 2016/679 of 27 April 2016 on the protection of natural persons with regard to the processing of personal data and on the free movement of such data, and repealing Directive 95/46/EC (General Data Protection Regulation).";
const MARKDOWN_INDENT = "    ";
/**
 * @name isBullet
 * @description
 * @param char
 * @returns
 */
function isBullet(char: string) {
  return char === "-";
}
/**
 * @name getLink
 * @description
 * @param link
 * @returns
 */
function getLink(link: string) {
  const matches = LINK_REGEXP.exec(link);
  if (matches && matches.length) return [matches[1], matches[2]];
  return [];
}
/**
 * @name calcWhiteSpaces
 * @description
 * @param string
 * @returns
 */
function calcWhiteSpaces(string: string) {
  const firstElement = string.split("-")[0];
  return firstElement.length;
}
/**
 * @name calcNestingLevel
 * @description
 * @param string
 * @returns
 */
function calcNestingLevel(string: string) {
  let idx = 0;
  if (!isBullet(string[0])) {
    const whiteSpaces = calcWhiteSpaces(string);
    idx = whiteSpaces / 4;
  }
  return Math.ceil(idx);
}
/**
 * @name getLastChild
 * @description
 * @param body
 * @returns
 */
function getLastChild(
  body: GoogleAppsScript.Document.Body,
): GoogleAppsScript.Document.Element {
  const childIndex = body.getNumChildren() - 1;
  return body.getChild(childIndex);
}
/**
 * @name getLastListItem
 * @description
 * @param body
 * @returns
 */
function getLastListItem(
  body: GoogleAppsScript.Document.Body,
): GoogleAppsScript.Document.Element | null {
  const lastChild = getLastChild(body);
  if (lastChild.getType() == DocumentApp.ElementType.LIST_ITEM) {
    return lastChild;
  }
  return null;
}
/**
 * @name deleteAllParagraphs
 * @description
 * @param body
 * @returns
 */
function deleteAllParagraphs(body: GoogleAppsScript.Document.Body) {
  body.insertParagraph(0, "");
  body.appendParagraph("");
  for (let i = body.getNumChildren(); i >= 0; i--) {
    try {
      body.getChild(i).removeFromParent();
    } catch (e) {
      Logger.log({ e });
    }
  }
}
/**
 * @name setListItem
 * @description
 * @param body
 * @param string
 * @param nestingLevel
 * @returns
 */
function setListItem(
  body: GoogleAppsScript.Document.Body,
  string: string,
  nestingLevel = 1,
): GoogleAppsScript.Document.ListItem {
  const listItem = body.appendListItem(`${string}`);
  listItem.setNestingLevel(nestingLevel);
  listItem.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
  return listItem;
}
/**
 * @name setIndent
 * @description
 * @param item
 * @param config
 * @returns
 */
function setIndent(
  item:
    | GoogleAppsScript.Document.Paragraph
    | GoogleAppsScript.Document.ListItem,
  indent = 0,
): GoogleAppsScript.Document.Element {
  item.setIndentFirstLine(indent);
  item.setIndentStart(indent);
  return item;
}
/**
 * @name setParagraph
 * @description
 * @param body
 * @param string
 * @param config
 * @returns
 */
function setParagraph(
  body: GoogleAppsScript.Document.Body,
  string: string,
  config?: {
    indent?: number;
  },
): GoogleAppsScript.Document.Paragraph {
  const idx = body.getNumChildren();
  const P = body.insertParagraph(idx, string);
  setIndent(P, config?.indent || 0);
  return P;
}
/**
 * @name getFormattedLocation
 * @description
 * @param location
 * @param locationType
 * @returns
 */
function getFormattedLocation(location: string, locationType: string) {
  return location
    ? `${location}${locationType ? `, ${locationType.toLowerCase()}` : ""}`
    : "";
}
/**
 * @name getFormattedIndustry
 * @description
 * @param industry
 * @returns
 */
function getFormattedIndustry(industry: string) {
  return industry ? industry.replace(/\d+ /, "") : "";
}
/**
 * @name getFormattedIsco8Code
 * @description
 * @param isco8Code
 * @returns
 */
function getFormattedIsco8Code(isco8Code: string) {
  return isco8Code ? isco8Code.replace(/\D+/g, "") : "";
}
/**
 * @name formatDate
 * @description
 * @param date
 * @returns
 */
function formatDate(date: Date) {
  const mm = date.getMonth() + 1;
  return [(mm > 9 ? "" : "0") + mm, date.getFullYear()].join("/");
}
/**
 * @name addLineBreak
 * @description
 * @param listItem
 * @param size
 * @returns
 */
function addLineBreak(
  listItem:
    | GoogleAppsScript.Document.Paragraph
    | GoogleAppsScript.Document.ListItem,
  size = 5,
) {
  return listItem.appendText(`\n `).setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: size,
  });
}
/**
 * @name getFormattedAt
 * @description
 * @param row
 * @returns
 */
function getFormattedAt(row: Row) {
  return row["start"] ? "at" : "issued by";
}
/**
 * @name getFormattedDatePrefix
 * @description
 * @param row
 * @returns
 */
function getFormattedDatePrefix(row: Row) {
  return !row["start"] ? "Issue date" : "Date";
}
/**
 * @name getFormattedTitle
 * @description
 * @param row
 * @returns
 */
function getFormattedTitle(row: Row) {
  let hours = "";
  if (row["hours"]) {
    const timeSpent = row["hours"];
    hours = ` (${timeSpent} ${timeSpent > 1 ? "hours" : "hour"})`;
  }
  return `${row.title}${hours}`;
}
/**
 * @name getFormattedDate
 * @description
 * @param row
 * @returns
 */
function getFormattedDate(row: Row) {
  let keys = [];
  if (row["start"]) {
    keys = ["start", "end"];
  } else {
    keys = ["issue date", "expiration date"];
  }
  const isEndKey = keys[1] === "end";
  const formattedStartDate = row[keys[0]] ? formatDate(row[keys[0]]) : "";
  const formattedEndDate = isEndKey
    ? row[keys[1]]
      ? ` - ${formatDate(row[keys[1]])}`
      : " - Present"
    : row[keys[1]]
    ? `, expires: ${formatDate(row[keys[1]])}`
    : "";
  const formattedTime = row?.time && isEndKey ? ` (${row.time})` : "";
  return `${formattedStartDate}${formattedEndDate}${formattedTime}`;
}
/**
 * @name isSection
 * @description
 * @param array
 * @returns
 */
function isSection(array: Row[]) {
  return Array.isArray(array) && array.filter((el) => el.enable).length > 1
    ? true
    : false;
}
/**
 * @name removeEmptyParagraph
 * @description
 * @param body
 * @returns
 */
function removeEmptyParagraph(body: GoogleAppsScript.Document.Body) {
  try {
    const child = body.getChild(0);
    body.removeChild(child);
    return body;
  } catch (e) {
    return body;
  }
}
/**
 * @name isDataSheet
 * @description
 * @param sheetName
 * @returns
 */
function isDataSheet(sheetName: string) {
  return (
    !sheetName.includes("Settings - ") &&
    !sheetName.includes("Range - ") &&
    !sheetName.includes("Resume") &&
    !sheetName.includes("Cover Letter")
  );
}

/**
 * @name filterCount
 * @description Filters count of same string in array
 * @param array of strings.
 * @param minCount
 * @return Array of occurrence objects sorted by occurrence and filtered based
 * on the minCount value
 */
function filterCount(
  array: string[],
  minCount: number,
): { value: string; count: number }[] {
  const obj: Record<string, number> = {};

  array.forEach((e) => (obj[e] = (obj[e] || 0) + 1));

  return Object.entries(obj)
    .map(([value, count]: [string, number]) => ({
      value,
      count,
    }))
    .sort((a, b) => (b.count as number) - (a.count as number))
    .filter((e) => (e.count as number) >= minCount);
}

/**
 * @name getSheetData
 * @description This function will construct an object using the header's cell value
 * or the header's note value
 * @param sheet A Google Sheet Object
 * @returns The databaseObjectArray.
 */
function getSheetData(sheet: GoogleAppsScript.Spreadsheet.Sheet): Row[] {
  const [headers, ...data] = sheet.getDataRange().getValues();

  const databaseObjectArray = data.map((row) => {
    return row.reduce((acc, value, i) => {
      const key = headers[i];
      if (key === "") return acc;
      return {
        ...acc,
        [key]: typeof value === "string" ? value.trim() : value,
      };
    }, {});
  });

  return databaseObjectArray;
}
/**
 * @name getAllSheetsData
 * @description
 * @returns The object of formatted sheet data.
 */
function getAllSheetsData(): Record<string, object[]> {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const data = {};

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();

    if (isDataSheet(sheetName)) {
      const sheetData = getSheetData(sheet);
      data[sheetName] = sheetData;
    }
  }
  return data;
}
/**
 * @name formatArray
 * @description
 * @param array
 * @param lang
 * @returns Formatted array
 */
function formatArray(
  array: {
    value: string;
    count: number;
  }[],
  lang = "en",
): string {
  const formatter = new Intl.ListFormat(lang, {
    style: "long",
    type: "conjunction",
  });
  return formatter.format(array.map((e) => e.value)).trim();
}
/**
 * @name makeId
 * @description
 * @see https://stackoverflow.com/questions/1349404/generate-random-string-characters-in-javascript
 * @param length
 * @returns string of random characters
 */
function makeId(length = 5) {
  let result = "";
  const characters =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  const charactersLength = characters.length;
  let counter = 0;
  while (counter < length) {
    result += characters.charAt(Math.floor(Math.random() * charactersLength));
    counter += 1;
  }
  return result;
}

type FileType = {
  doc: GoogleAppsScript.Document.Document;
  body: GoogleAppsScript.Document.Body;
  id: string;
  markdown: string[];
  current:
    | GoogleAppsScript.Document.Paragraph
    | GoogleAppsScript.Document.ListItem;
};
/**
 * @name createFiles
 * @description
 * @returns
 */
function createFiles() {
  const folder = DriveApp.createFolder("Resumer folder");
  const folderId = folder.getId();
  const details = ["resume", "cover_letter"].reduce<
    Record<"resume" | "cover_letter", FileType> & { folderId: string }
  >(
    (acc, val, idx) => {
      const doc = DocumentApp.create(`Resumer - ${idx}`);
      const id = doc.getId();
      DriveApp.getFileById(id).moveTo(folder);
      const body = doc.getBody();
      const markdown = [];
      const current = "";

      addResume("", "");
      body.setText("");
      deleteAllParagraphs(body);
      body.clear();

      acc[val] = {
        doc,
        body,
        id,
        markdown,
        current,
      };

      return acc;
    },
    { resume: {}, cover_letter: {}, folderId } as {
      resume: FileType;
      cover_letter: FileType;
      folderId: string;
    },
  );

  return details;
}
