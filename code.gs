// Original code written by Matthews Ma August 30, 2018
// Edits by Levi Moore August 31, 2018

/* Setup and global variables */

var responsesSheet = "Responses"; // The spreadsheet with the from responses. The sheet should be the first one.
var emailSheet = "Emails";
var schoolCol = 6; // The column where the school name is. 1 = A, 2 = B, etc...
var emailCol = 9; // The column where the email address is.
var takeCol = 2;
var dateCol = 3; // The column where the date is.
var timeCol = 4; // The column where the length of the program is. Full, AM, PM.
var programCol = 11;
var nameCol = 7;
var centreCol = 5;

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var formResponses = spreadsheet.getSheetByName(responsesSheet);
var emailList = spreadsheet.getSheetByName(emailSheet);

var centres = ["Blair", "Laurel", "Heidelberg", "Huron", "Wrigley"];


/* UI */

function onOpen() {
    // Create a custom menu.
    // Google scripts only supports creating menus manually. Menu items also cannot
    // pass arguments to functions. Therefore, the items and functions must be hardcoded.
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Scripts")
        .addItem("Sort Entries", "sortEntries")
        .addItem("Draft Emails", "draftEmails")
        .addSubMenu(ui.createMenu("Sync Calendar")
            .addItem("Sync Blair Calendar", "syncCalendarBlair")
            .addItem("Sync Laurel Calendar", "syncCalendarLaurel")
            .addItem("Sync Heidelberg Calendar", "syncCalendarHeidelberg")
            .addItem("Sync Huron Calendar", "syncCalendarHuron")
            .addItem("Sync Wrigley Calendar", "syncCalendarWrigley"))
        .addToUi();
}


function removeEmptyRows() {
    // Sorting and creating new spreadsheets quickly surpasses Google's
    // 2 million cell limit. This function removes empty rows from all
    // sheets.
    for (i = 0; i < spreadsheet.getSheets().length; i++) {
        var sheet = spreadsheet.getSheets()[i];
        var maxRows = sheet.getMaxRows();
        var lastRow = sheet.getLastRow();
        if (maxRows - lastRow != 0) {
            sheet.deleteRows(lastRow + 1, maxRows - lastRow);
        }
    }
}


function numberOfCells() {
    // Returns the total number of cells in all spreadsheets.
    var formatThousandsNoRounding = function (n, dp) {
        var e = '', s = e + n, l = s.length, b = n < 0 ? 1 : 0,
            i = s.lastIndexOf('.'), j = i == -1 ? l : i,
            r = e, d = s.substr(j + 1, dp);
        while ((j -= 3) > b) {
            r = ',' + s.substr(j, 3) + r;
        }
        return s.substr(0, j + 3) + r +
            (dp ? '.' + d + (d.length < dp ?
                ('00000').substr(0, dp - d.length) : e) : e);
    };
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    var cells_count = 0;
    for (var i in sheets) {
        cells_count += (sheets[i].getMaxColumns() * sheets[i].getMaxRows());
    }
    Logger.log(formatThousandsNoRounding(cells_count))
}


/* Sorting */

function sortEntries() {
    // Sorts entries by school into individual spreadsheets.

    // Check if the entry's school has a spreadsheet. If not, create one.
    for (i = 1; i < formResponses.getLastRow(); i++) {
        var schoolName = formResponses.getRange(i + 1, schoolCol).getValue();
        var take = formResponses.getRange(i + 1, takeCol).getValue();

        if (spreadsheet.getSheetByName(schoolName) != null) {
            continue;
        } else {
            // Create the school's spreadsheet and setup with a query statement.
            var schoolSheet = spreadsheet.insertSheet(schoolName);

            var query = '=query(' + responsesSheet + '!1:1000, "select B, C, D, E, F, G, H, J, I where F=\'' +
                schoolName + '\'", -1)';
            schoolSheet.getRange(1, 1).setValue(query);
        }
    }
}


/* Emails */

function draftEmails() {
    // Sends an email to all schools with a pdf of their sheet attached.

    function findEmail() {
        var data = emailList.getDataRange().getValues();
        for (i = 0; i < data.length; i++) {
            if (data[i][0] == schoolName) {
                Logger.log("run" + i);
                var principalEmail = emailList.getRange(i + 1, 2).getValue();
                return principalEmail;
            }
        }
    }

    for (i = 1; i < spreadsheet.getSheets().length - 1; i++) {
        var sheet = spreadsheet.getSheets()[i]; // Sends the email to the first entry in the school spreadsheet.
        var schoolName = sheet.getName();
        var teacherEmail = sheet.getRange(2, emailCol).getValue();

        var principalEmail = findEmail();

        // Customize the email to your liking.
        var pdfName = "OEE Schedule for " + schoolName + ".";
        var emails = [teacherEmail, principalEmail];
        var subject = "OEE Schedule for " + schoolName + ".";
        var body = "Hello!\nThank you for registering for the OEE program. Please see attached for your school's schedule.";
        draftPdfEmail(i, pdfName, emails, subject, body);
        Logger.log(schoolName + teacherEmail + spreadsheet.getSheets().length);
    }
}


function draftPdfEmail(sheetNumber, pdfName, emails, subject, body) {
    // Convert a spreadsheet to a pdf, attach it, and send it as an email.
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = spreadsheet.getId();
    var sheetId = sheetNumber ? spreadsheet.getSheets()[sheetNumber].getSheetId() : null;
    var url_base = spreadsheet.getUrl().replace(/edit$/, '');

    var url_ext = 'export?exportFormat=pdf&format=pdf'   // Export as PDF.

        + (sheetId ? ('&gid=' + sheetId) : ('&id=' + spreadsheetId))
        // Parameters
        + '&size=A4'          // Paper size
        + '&portrait=true'    // Orientation, false for landscape
        + '&fitw=true'        // Fit to width, false for actual size
        + '&sheetnames=true&printtitle=false&pagenumbers=true'  // Hide optional headers and footers
        + '&gridlines=true'  // false = Hide gridlines, true = show gridlines
        + '&fzr=false';       // Do not repeat row headers (frozen rows) on each page

    var options = {
        headers: {
            'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        }
    }

    var response = UrlFetchApp.fetch(url_base + url_ext, options);
    var blob = response.getBlob().setName(pdfName + '.pdf');
    if (emails) {
        var mailOptions = {
            attachments: blob
        }
        GmailApp.createDraft(emails, subject, body, mailOptions);
    }
}


/* Calendar */

// Wrapper functions
function syncCalendarBlair() {
    syncCalendar("Blair");
}

function syncCalendarLaurel() {
    syncCalendar("Laurel");
}

function syncCalendarHeidelberg() {
    syncCalendar("Heidelberg");
}

function syncCalendaryHuron() {
    syncCalendar("Huron");
}

function syncCalendarWrigley() {
    syncCalendar("Wrigley");
}

// Main function
function syncCalendar(c) {
    // Resets centre calendar on call.
    if (CalendarApp.getCalendarsByName(c)[0] != null) {
        CalendarApp.getCalendarsByName(c)[0].deleteCalendar();
    }

    CalendarApp.createCalendar(c);
    Logger.log(c);

    for (i = 1; i < spreadsheet.getSheetByName(responsesSheet).getLastRow(); i++) {
        var centre = formResponses.getRange(i + 1, centreCol).getValue();

        if (formResponses.getRange(i + 1, takeCol).getValue() == "x" || centre != c) {
            continue;
        }

        var date = formResponses.getRange(i + 1, dateCol).getValue();
        var time = formResponses.getRange(i + 1, timeCol).getValue();
        var program = formResponses.getRange(i + 1, programCol).getValue();
        var teacherName = formResponses.getRange(i + 1, nameCol, 1, 2).getValues();
        var startTime = new Date(date);
        var endTime = new Date(date);

        var title = program + " with " + teacherName[0][0] + " " + teacherName[0][1];
        if (time == "Full") {
            startTime.setHours(9);
            endTime.setHours(15);
        } else if (time == "AM") {
            startTime.setHours(9);
            endTime.setHours(12);
        } else if (time == "PM") {
            startTime.setHours(12);
            endTime.setHours(15);
        } else {
            continue;
        }

        CalendarApp.getCalendarsByName(centre)[0].createEvent(title, startTime, endTime)
    }
}
