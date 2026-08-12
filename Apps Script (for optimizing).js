// ==========================================================================
// UNIVERSAL HELPER: Case-Insensitive User Matching
// ==========================================================================
function isMatchingUser(sheetValue, inputUser) {
  if (!sheetValue || !inputUser) return false;
  return sheetValue.toString().trim().toLowerCase() === inputUser.toString().trim().toLowerCase();
}

// ==========================================================================
// MAIN GET HANDLER
// ==========================================================================
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users") || ss.getSheets()[0];

  // Load Discover & Explore Game States
  if (e.parameter.action === 'loadDiscover') return handleLoadDiscover(e.parameter.user);
  if (e.parameter.action === 'loadExplore') return handleLoadExplore(e.parameter.user);

  // ACTION: Get Roster (Optimized: Selective Column Reading & Avatar Length Guard)
  if (e.parameter.action === "getRoster") {
    var sheet = ss.getSheetByName("Users");
    var lastRow = sheet.getLastRow();
    var users = [];

    if (lastRow > 1) {
      // 1. Only read Columns A to E (Row 2, Col 1, NumRows, 5 Columns)
      // This completely avoids loading heavy JSON strings in Column G (flags)
      var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

      for (var i = 0; i < data.length; i++) {
        var name = data[i][1];
        
        // Ensure row actually contains a student name
        if (name && String(name).trim() !== "") {
          var avatar = data[i][4] || "new_user.png";

          // 2. Prevent Base64 data-URL bloat from crashing the HTTP payload
          if (String(avatar).length > 200) {
            avatar = "new_user.png";
          }

          users.push({
            name: String(name).trim(),
            avatar: String(avatar).trim(),
            // Compact epoch timestamp for smaller JSON transmission
            timestamp: data[i][0] ? new Date(data[i][0]).getTime() : 0
          });
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      users: users
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // ACTION 2: Get a Specific User's Data (Avatar, Total Combined Points & Flags)
  if (e.parameter.action === 'getUserData') {
    var user = e.parameter.user;
    if (!user) return ContentService.createTextOutput(JSON.stringify({ error: "No parameters provided" })).setMimeType(ContentService.MimeType.JSON);

    var data = sheet.getDataRange().getValues();
    var basePoints = 0;
    var avatar = "new_user.png";
    var flags = "{}";
    var userFound = false;

    // 1. Find user in the main Users sheet
    for (var i = data.length - 1; i >= 1; i--) {
      if (isMatchingUser(data[i][1], user)) {
        userFound = true;
        avatar = data[i][4] || "new_user.png";
        basePoints = Number(data[i][5]) || 0; // Column F
        flags = data[i][6] || "{}"; // Column G
        break;
      }
    }

    if (userFound) {
      var discoverPoints = 0;
      var explorePoints = 0;

      // 2. Fetch Discover Points (Only query if data rows exist)
      var discoverSheet = ss.getSheetByName("Discover");
      if (discoverSheet && discoverSheet.getLastRow() > 1) {
        var dData = discoverSheet.getDataRange().getValues();
        for (var j = dData.length - 1; j >= 1; j--) {
          if (isMatchingUser(dData[j][0], user)) {
            discoverPoints = Number(dData[j][1]) || 0; // Column B (Index 1)
            break;
          }
        }
      }

      // 3. Fetch Explore Points (Only query if data rows exist)
      var exploreSheet = ss.getSheetByName("Explore");
      if (exploreSheet && exploreSheet.getLastRow() > 1) {
        var eData = exploreSheet.getDataRange().getValues();
        for (var k = eData.length - 1; k >= 1; k--) {
          if (isMatchingUser(eData[k][0], user)) {
            explorePoints = Number(eData[k][1]) || 0; // Column B (Index 1)
            break;
          }
        }
      }

      // 4. Calculate Final Sum
      var totalPoints = basePoints + discoverPoints + explorePoints;

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        avatar: avatar,
        points: totalPoints,
        flags: flags
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, points: 0, flags: "{}" })).setMimeType(ContentService.MimeType.JSON);
  }

  // ACTION 3: Get all Reflections for the Community Board
  if (e.parameter.action === 'getReflections') {
    var reflectSheet = ss.getSheetByName("Reflections");
    if (!reflectSheet || reflectSheet.getLastRow() <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ success: true, reflections: [] })).setMimeType(ContentService.MimeType.JSON);
    }

    var data = reflectSheet.getDataRange().getValues();
    var reflections = [];

    for (var i = data.length - 1; i >= 1; i--) {
      reflections.push({
        timestamp: data[i][0], user: data[i][1], avatar: data[i][2],
        idea: data[i][3], prob: data[i][4], people: data[i][5],
        sol: data[i][6], money: data[i][7], res: data[i][8], step: data[i][9]
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true, reflections: reflections }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ACTION 4: Get Password, Flags, Total Combined Points & Avatar on Login
  var userParam = e.parameter.user;
  if (userParam) {
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (isMatchingUser(data[i][1], userParam)) {
        var basePoints = Number(data[i][5]) || 0;
        var avatar = data[i][4] || "new_user.png";
        var discoverPoints = 0;
        var explorePoints = 0;

        // Fetch Discover Points
        var discoverSheet = ss.getSheetByName("Discover");
        if (discoverSheet && discoverSheet.getLastRow() > 1) {
          var dData = discoverSheet.getDataRange().getValues();
          for (var j = dData.length - 1; j >= 1; j--) {
            if (isMatchingUser(dData[j][0], userParam)) {
              discoverPoints = Number(dData[j][1]) || 0;
              break;
            }
          }
        }

        // Fetch Explore Points
        var exploreSheet = ss.getSheetByName("Explore");
        if (exploreSheet && exploreSheet.getLastRow() > 1) {
          var eData = exploreSheet.getDataRange().getValues();
          for (var k = eData.length - 1; k >= 1; k--) {
            if (isMatchingUser(eData[k][0], userParam)) {
              explorePoints = Number(eData[k][1]) || 0;
              break;
            }
          }
        }

        var totalPoints = basePoints + discoverPoints + explorePoints;

        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          sequence: data[i][3],
          flags: data[i][6] || "{}",
          points: totalPoints,
          avatar: avatar
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ error: "User not found" })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================================================
// MAIN POST HANDLER
// ==========================================================================
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users") || ss.getSheets()[0];

    logToSpreadsheet("TRAFFIC_IN", "Params: " + JSON.stringify(e.parameter) + " | PostData: " + (e.postData ? e.postData.contents : "none"));

    var data = null;
    if (e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      throw new Error("No readable data payload found.");
    }

    // =========================================
    // ACTION ROUTING
    // =========================================

    if (data.action === 'saveExplore') return handleSaveExplore(data);
    if (data.action === 'saveDiscover') return handleSaveDiscover(data);

    // ACTION: Save User Flags
    if (data.action === 'saveUserFlags') {
      var rows = sheet.getDataRange().getValues();
      for (var i = rows.length - 1; i >= 1; i--) {
        if (isMatchingUser(rows[i][1], data.user)) {
          var existingFlagsStr = rows[i][6] || "{}";
          var existingFlags = {};

          try { existingFlags = JSON.parse(existingFlagsStr); } catch (err) { }
          if (typeof existingFlags !== 'object' || existingFlags === null) existingFlags = {};

          var updatedFlags = Object.assign(existingFlags, data.flags || {});

          // Ensure Column G has a header before saving
          if (sheet.getRange("G1").getValue() === "") {
            sheet.getRange("G1").setValue("Flags JSON");
          }

          sheet.getRange(i + 1, 7).setValue(JSON.stringify(updatedFlags)); // Save to Column G
          return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: "User not found" })).setMimeType(ContentService.MimeType.JSON);
    }

    // ACTION: Submit a Reflection Form & Award Points
    if (data.action === 'submitReflection') {
      var reflectSheet = ss.getSheetByName("Reflections");

      if (!reflectSheet) {
        reflectSheet = ss.insertSheet("Reflections");
        reflectSheet.appendRow(["Timestamp", "User", "Avatar", "Idea", "Problem", "People", "Solution", "Money", "Resources", "Step"]);
        reflectSheet.setFrozenRows(1);
      }

      var alreadySubmitted = false;
      var reflectData = reflectSheet.getDataRange().getValues();
      for (var j = reflectData.length - 1; j >= 1; j--) {
        if (isMatchingUser(reflectData[j][1], data.user)) {
          alreadySubmitted = true;
          break;
        }
      }

      reflectSheet.appendRow([new Date(), data.user, data.avatar, data.idea, data.prob, data.people, data.sol, data.money, data.res, data.step]);

      if (!alreadySubmitted) {
        if (sheet.getRange("F1").getValue() === "") sheet.getRange("F1").setValue("Points");

        var rows = sheet.getDataRange().getValues();
        for (var i = rows.length - 1; i >= 1; i--) {
          if (isMatchingUser(rows[i][1], data.user)) {
            var currentPoints = Number(rows[i][5]) || 0;
            sheet.getRange(i + 1, 6).setValue(currentPoints + 1000);
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }

    // ACTION: Update Avatar
    if (data.action === 'updateAvatar') {
      var rows = sheet.getDataRange().getValues();
      for (var i = rows.length - 1; i >= 1; i--) {
        if (isMatchingUser(rows[i][1], data.fullName)) {
          sheet.getRange(i + 1, 5).setValue(data.avatar);
          return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: "User not found" })).setMimeType(ContentService.MimeType.JSON);
    }

    // =========================================
    // DEFAULT ACTION: Registration
    // =========================================
    if (!data.action || data.action === 'register') {
      if (!data.pictographicPassword) {
        throw new Error("Missing pictographicPassword in registration payload. Payload received: " + JSON.stringify(data));
      }
      sheet.appendRow([data.timestamp, data.fullName, data.dob, data.pictographicPassword.join(" -> "), "new_user.png", 0, "{}"]);
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }

    throw new Error("Unrecognized action received: " + data.action + " | Full payload: " + JSON.stringify(data));

  } catch (globalError) {
    logToSpreadsheet("FATAL_CRASH", globalError.toString() + "\nStack: " + globalError.stack);
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: globalError.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================================================
// EXPLORE MODULE SAVE & LOAD HANDLERS
// ==========================================================================
function handleLoadExplore(user) {
  if (!user) return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'No user provided' })).setMimeType(ContentService.MimeType.JSON);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Explore");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: 'not_found' })).setMimeType(ContentService.MimeType.JSON);

  var dataRange = sheet.getDataRange().getValues();
  for (var i = 1; i < dataRange.length; i++) {
    if (isMatchingUser(dataRange[i][0], user)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: dataRange[i][5] })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'not_found' })).setMimeType(ContentService.MimeType.JSON);
}

function handleSaveExplore(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Explore");

  if (!sheet) {
    sheet = ss.insertSheet("Explore");
    sheet.appendRow(["Full Name", "Total Points", "Points Breakdown", "Ending Reached", "Self Review", "Save Data JSON", "Last Updated"]);
    sheet.setFrozenRows(1);
  }

  var fullName = data.user;
  var ending = data.ending || "N/A";
  var review = data.review ? JSON.stringify(data.review) : "N/A";
  var timestamp = new Date();

  var totalPoints = 0;
  var pointsBreakdown = "{}";

  if (data.points) {
    if (typeof data.points === 'object') {
      pointsBreakdown = JSON.stringify(data.points);
      for (var key in data.points) {
        if (typeof data.points[key] === 'number') {
          totalPoints += data.points[key];
        }
      }
    } else {
      totalPoints = Number(data.points) || 0;
      pointsBreakdown = data.points.toString();
    }
  }

  var dataRange = sheet.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < dataRange.length; i++) {
    if (isMatchingUser(dataRange[i][0], fullName)) {
      rowIndex = i + 1;
      break;
    }
  }

  // Batched 7-column write
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 7).setValues([[fullName, totalPoints, pointsBreakdown, ending, review, data.saveData, timestamp]]);
  } else {
    sheet.appendRow([fullName, totalPoints, pointsBreakdown, ending, review, data.saveData, timestamp]);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================================================
// DISCOVER MODULE SAVE & LOAD HANDLERS
// ==========================================================================
function handleLoadDiscover(user) {
  if (!user) return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'No user provided' })).setMimeType(ContentService.MimeType.JSON);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Discover");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: 'not_found' })).setMimeType(ContentService.MimeType.JSON);

  var dataRange = sheet.getDataRange().getValues();
  for (var i = 1; i < dataRange.length; i++) {
    if (isMatchingUser(dataRange[i][0], user)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: dataRange[i][5] })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'not_found' })).setMimeType(ContentService.MimeType.JSON);
}

function handleSaveDiscover(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Discover");

  if (!sheet) {
    sheet = ss.insertSheet("Discover");
    sheet.appendRow(["Full Name", "Total Points", "Total Coins", "Points Breakdown", "Coins Breakdown", "Save Data JSON", "Last Updated"]);
    sheet.setFrozenRows(1);
  }

  var fullName = data.user;
  var timestamp = new Date();
  var dataRange = sheet.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < dataRange.length; i++) {
    if (isMatchingUser(dataRange[i][0], fullName)) {
      rowIndex = i + 1;
      break;
    }
  }

  // Batched 7-column write
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 7).setValues([[fullName, data.points || 0, data.coins || 0, data.categoryPoints || "{}", data.categoryCoins || "{}", data.saveData, timestamp]]);
  } else {
    sheet.appendRow([fullName, data.points || 0, data.coins || 0, data.categoryPoints || "{}", data.categoryCoins || "{}", data.saveData, timestamp]);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================================================
// ERROR & EVENT LOGGER
// ==========================================================================
function logToSpreadsheet(type, message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("ScriptLogs");
    if (!logSheet) {
      logSheet = ss.insertSheet("ScriptLogs");
      logSheet.appendRow(["Timestamp", "Log Type", "Details / Error Message"]);
      logSheet.setFrozenRows(1);
    }
    logSheet.appendRow([new Date(), type, message]);
  } catch (e) { }
}