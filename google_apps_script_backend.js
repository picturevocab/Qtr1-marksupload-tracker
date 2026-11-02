/**
 * QUARTER 1 MARKS UPLOAD TRACKER - GOOGLE APPS SCRIPT BACKEND (WITH USER LOGIN)
 * -------------------------------------------------------------
 * Setup Instructions:
 * 1. Create a Google Sheet named "MarksTracker".
 *    Headers: [Grade, Section, Program, StudentCount, Status, CompletedMarks]
 * 2. Create a Google Sheet tab named "Users" for logins (Auto-created if missing).
 *    Headers: [UserID, Password, Name, Role]
 * 3. Deploy -> Web App -> Execute as: Me -> Access: Anyone.
 */

const SHEET_NAME = "MarksTracker";
const USERS_SHEET_NAME = "Users";

/**
 * Add custom menu to Google Sheets for easy initialization
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Marks Tracker')
    .addItem('Populate/Reset All 75 Sections', 'initializeSheetStructure')
    .addToUi();
}

/**
 * Utility to standardise JSON output
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper to validate user credentials against the Users sheet
 */
function validateUser(userId, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let userSheet = ss.getSheetByName(USERS_SHEET_NAME);

  // Auto-create default users table if missing
  if (!userSheet) {
    userSheet = ss.insertSheet(USERS_SHEET_NAME);
    userSheet.appendRow(["UserID", "Password", "Name", "Role"]);
    userSheet.appendRow(["admin", "2358", "Administrator", "Admin"]);
  }

  const usersData = userSheet.getDataRange().getValues();
  for (let i = 1; i < usersData.length; i++) {
    const row = usersData[i];
    if (String(row[0]).toLowerCase() === String(userId).toLowerCase()) {
      // If password is provided, check password match
      if (password !== undefined) {
        if (String(row[1]) === String(password)) {
          return { valid: true, userId: row[0], name: row[2], role: row[3] };
        }
      } else {
        // If checking session/token validity only
        return { valid: true, userId: row[0], name: row[2], role: row[3] };
      }
    }
  }
  return { valid: false };
}

/**
 * Force creates or resets the MarksTracker sheet with all 75 sections
 */
function initializeSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Header Row
  sheet.appendRow(["Grade", "Section", "Program", "StudentCount", "Status", "CompletedMarks"]);

  const sectionsList = [
    [4, "4A"], [4, "4B"], [4, "4C"], [4, "4D"], [4, "4E"], [4, "4F"], [4, "4G"], [4, "4H"], [4, "4I"], [4, "4J"], [4, "4K"], [4, "4L"],
    [5, "5A"], [5, "5B"], [5, "5C"], [5, "5D"], [5, "5E"], [5, "5F"], [5, "5G"], [5, "5H"], [5, "5I"], [5, "5J"], [5, "5K"],
    [6, "6A"], [6, "6B"], [6, "6C"], [6, "6D"], [6, "6E"], [6, "6F"], [6, "6G"], [6, "6H"],
    [8, "8A"], [8, "8B"], [8, "8C"], [8, "8D"], [8, "8E"], [8, "8F"], [8, "8G"], [8, "8H"], [8, "8I"], [8, "8J"], [8, "8K"], [8, "8L"],
    [9, "9A"], [9, "9B"], [9, "9C"], [9, "9D"], [9, "9E"], [9, "9F"], [9, "9G"], [9, "9H"], [9, "9I"], [9, "9J"], [9, "9K"],
    [10, "10A"], [10, "10B"], [10, "10C"], [10, "10D"], [10, "10E"], [10, "10F"], [10, "10G"], [10, "10H"], [10, "10I"], [10, "10J"],
    [11, "11A"], [11, "11B"], [11, "11D"], [11, "11E"], [11, "11F"], [11, "11G"], [11, "11H"], [11, "11I"],
    [12, "12A"], [12, "12B"], [12, "12C"], [12, "12D"], [12, "12E"], [12, "12F"], [12, "12G"], [12, "12H"], [12, "12I"]
  ];

  const mockRows = sectionsList.map(([grade, sec]) => {
    const program = grade >= 9 ? "IGCSE Program" : "Normal Program";
    return [grade, sec, program, 25, "Pending", 0];
  });

  mockRows.forEach(row => sheet.appendRow(row));
  return sheet;
}

/**
 * Handle GET requests - Fetches marks data or performs login verification
 */
function doGet(e) {
  try {
    const params = e ? e.parameter : {};
    const action = params.action || 'getData';
    const userId = params.user_id;
    const password = params.password;
    const callback = params.callback;

    let responseData;

    // Handle Login Action
    if (action === 'login') {
      if (!userId || !password) {
        responseData = { success: false, error: "Missing user_id or password." };
      } else {
        const auth = validateUser(userId, password);
        if (auth.valid) {
          responseData = {
            success: true,
            message: "Login successful",
            user: { userId: auth.userId, name: auth.name, role: auth.role }
          };
        } else {
          responseData = { success: false, error: "Invalid User ID or Password." };
        }
      }
    } else {
      // Authenticate user session for reading data
      if (userId) {
        const authCheck = validateUser(userId);
        if (!authCheck.valid) {
          responseData = { success: false, error: "Unauthorized access. Invalid user_id." };
        }
      }

      if (!responseData) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(SHEET_NAME);
        
        // Auto-create full 75 sections structure if missing or if only sample data exists (<10 rows)
        if (!sheet || sheet.getLastRow() < 10) {
          sheet = initializeSheetStructure();
        }

        const data = sheet.getDataRange().getValues();
        const rows = data.slice(1);

        const result = rows.map((row, index) => ({
          rowIndex: index + 2,
          grade: row[0],
          section: String(row[1]),
          program: String(row[2]),
          studentCount: Number(row[3]) || 0,
          status: String(row[4] || "Pending"),
          completedMarks: Number(row[5]) || 0
        }));

        responseData = {
          success: true,
          data: result
        };
      }
    }

    if (callback) {
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(responseData) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return createJsonResponse(responseData);

  } catch (error) {
    const errorData = {
      success: false,
      error: error.toString()
    };
    const params = e ? e.parameter : {};
    const callback = params.callback;
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(errorData) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return createJsonResponse(errorData);
  }
}

/**
 * Handle POST requests - Updates section status
 */
function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const { user_id, rowIndex, section, status } = contents;

    // Verify User Authorization
    if (!user_id) {
      return createJsonResponse({ success: false, error: "Missing user_id in request payload." });
    }
    const authCheck = validateUser(user_id);
    if (!authCheck.valid) {
      return createJsonResponse({ success: false, error: "Unauthorized operation. Invalid user_id." });
    }

    // Restrict Viewer roles from modifying data
    if (authCheck.role && String(authCheck.role).toLowerCase() === "viewer") {
      return createJsonResponse({ success: false, error: "Access Denied: Viewers do not have editing privileges." });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error("Sheet MarksTracker not found.");
    }

    let targetRow = rowIndex;

    // Search by section name if rowIndex is omitted
    if (!targetRow && section) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]).toLowerCase() === String(section).toLowerCase()) {
          targetRow = i + 1;
          break;
        }
      }
    }

    if (!targetRow) {
      throw new Error("Target row or section not found.");
    }

    // Update Status (Column E = 5)
    if (status !== undefined) {
      sheet.getRange(targetRow, 5).setValue(status);
      
      // Update CompletedMarks (Column F = 6) based on Status
      const studentCount = sheet.getRange(targetRow, 4).getValue() || 25;
      if (status === 'Completed') {
        sheet.getRange(targetRow, 6).setValue(studentCount);
      } else if (status === 'Pending') {
        sheet.getRange(targetRow, 6).setValue(0);
      }
    }

    const updatedRow = sheet.getRange(targetRow, 1, 1, 6).getValues()[0];

    return createJsonResponse({
      success: true,
      message: "Update successful",
      updatedBy: authCheck.name,
      data: {
        rowIndex: targetRow,
        grade: updatedRow[0],
        section: String(updatedRow[1]),
        program: String(updatedRow[2]),
        studentCount: Number(updatedRow[3]) || 0,
        status: String(updatedRow[4]),
        completedMarks: Number(updatedRow[5]) || 0
      }
    });

  } catch (error) {
    return createJsonResponse({
      success: false,
      error: error.toString()
    });
  }
}