function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Cotton Club BD')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMachineListHTML() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  var options = '<option value="">-- Select --</option>';
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    initSettings(sheet);
  }
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var data = sheet.getRange(2, 1, lastRow, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() !== "") {
      options += '<option value="' + data[i][0].toString().trim() + '">' + data[i][0].toString().trim() + '</option>';
    }
  }
  return options;
}

function initSettings(sheet) {
  sheet.getRange("A1").setValue("Machine List");
  
  // Machine 87-112 + 137 + 222 + 223
  var machines = [];
  for (var i = 87; i <= 112; i++) {
    machines.push([String(i)]);
  }
  machines.push(["137"]);
  machines.push(["222"]);
  machines.push(["223"]);
  sheet.getRange(2, 1, machines.length, 1).setValues(machines);
  
  sheet.getRange("C1").setValue("Username");
  sheet.getRange("D1").setValue("Password");
  sheet.getRange("E1").setValue("Status");
  sheet.getRange("F1").setValue("CanViewOthers");
  sheet.getRange("G1").setValue("Permission");
  sheet.getRange("H1").setValue("AssignedShift");
  sheet.getRange("C2").setValue("admin");
  sheet.getRange("D2").setValue("1234");
  sheet.getRange("E2").setValue("Active");
  sheet.getRange("F2").setValue("Yes");
  sheet.getRange("G2").setValue("CanEdit");
  sheet.getRange("H2").setValue("All");
  SpreadsheetApp.flush();
}

function checkLogin(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return { success: false, message: "Settings not found" };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "No users found" };
  var data = sheet.getRange(2, 3, lastRow, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    var u = (data[i][0] || "").toString().trim();
    var p = (data[i][1] || "").toString().trim();
    var status = (data[i][2] || "Active").toString().trim();
    var canView = (data[i][3] || "No").toString().trim();
    var permission = (data[i][4] || "ViewOnly").toString().trim();
    var assignedShift = (data[i][5] || "A").toString().trim();
    if (u === username && p === password && u !== "") {
      if (status.toLowerCase() === "inactive") return { success: false, message: "Account is disabled. Contact Admin." };
      if (status.toLowerCase() === "pending") return { success: false, message: "Account pending approval. Wait for Admin." };
      return {
        success: true,
        username: u,
        isAdmin: (u.toLowerCase() === "admin"),
        canViewOthers: (canView.toLowerCase() === "yes" || u.toLowerCase() === "admin"),
        permission: (u.toLowerCase() === "admin") ? "CanEdit" : permission,
        assignedShift: (u.toLowerCase() === "admin") ? "All" : assignedShift
      };
    }
  }
  return { success: false, message: "Username or Password is incorrect" };
}

function registerUser(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    initSettings(sheet);
  }
  username = (username || "").toString().trim();
  password = (password || "").toString().trim();
  if (!username || !password) return "Username and Password required";
  if (username.toLowerCase() === "admin") return "Cannot register as admin";
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 3, Math.max(lastRow, 2), 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim().toLowerCase() === username.toLowerCase()) {
      return "Username already exists";
    }
  }
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 3).setValue(username);
  sheet.getRange(newRow, 4).setValue(password);
  sheet.getRange(newRow, 5).setValue("Pending");
  sheet.getRange(newRow, 6).setValue("No");
  sheet.getRange(newRow, 7).setValue("ViewOnly");
  sheet.getRange(newRow, 8).setValue("A");
  return "OK";
}

function getAllUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 3, lastRow, 6).getValues();
  var users = [];
  for (var i = 0; i < data.length; i++) {
    var u = (data[i][0] || "").toString().trim();
    if (u !== "") {
      users.push({
        username: u,
        password: (data[i][1] || "").toString(),
        status: (data[i][2] || "Active").toString(),
        canViewOthers: (data[i][3] || "No").toString(),
        permission: (data[i][4] || "ViewOnly").toString(),
        assignedShift: (data[i][5] || "A").toString()
      });
    }
  }
  return users;
}

function saveUser(userData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return "Settings sheet not found";
  if (userData.username.toLowerCase() === "admin") {
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 3, lastRow, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if ((data[i][0] || "").toString().trim().toLowerCase() === "admin") {
        sheet.getRange(i + 2, 4).setValue(userData.password);
        return "Admin password updated!";
      }
    }
    return "Admin not found";
  }
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 3, Math.max(lastRow, 2), 1).getValues();
  var foundRow = -1;
  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim() === userData.username) {
      foundRow = i + 2;
      break;
    }
  }
  if (foundRow === -1) {
    foundRow = sheet.getLastRow() + 1;
    sheet.getRange(foundRow, 3).setValue(userData.username);
  }
  sheet.getRange(foundRow, 4).setValue(userData.password);
  sheet.getRange(foundRow, 5).setValue(userData.status || "Active");
  sheet.getRange(foundRow, 6).setValue(userData.canViewOthers || "No");
  sheet.getRange(foundRow, 7).setValue(userData.permission || "ViewOnly");
  sheet.getRange(foundRow, 8).setValue(userData.assignedShift || "A");
  return "User saved successfully!";
}

function deleteUser(username) {
  if (username.toLowerCase() === "admin") return "Cannot delete admin!";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return "Sheet not found";
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 3, lastRow, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim() === username) {
      sheet.getRange(i + 2, 3, 1, 6).clearContent();
      return "User deleted!";
    }
  }
  return "User not found";
}

function getMachineData(machineNo, selectedDate, shift, currentUser, canViewOthers, permission, assignedShift) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Database');
  const emptyRolls = Array(10).fill({val: "-", bg: "#ffffff", strike: "none", color: "#000000"});
  if (!sheet || sheet.getLastRow() < 2) {
    return { nextRoll: 'R1', totalWeight: 0, todayTotal: 0, lastProgram: "", fullProgramStr: "", rollData: emptyRolls, enteredBy: "-", programBreakdown: [] };
  }
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow, 16).getValues();
  let nextRoll = 1;
  let rolls = emptyRolls.slice();
  let totalWeight = 0;
  let todayTotalAll = 0;
  let lastProgram = "";
  let fullProgStr = "";
  let enteredBy = "-";
  let targetRowIndex = -1;
  let isAdmin = (currentUser || "").toLowerCase() === "admin";
  let programBreakdown = [];

  for (let i = 0; i < data.length; i++) {
    let rDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    let rMachine = (data[i][1] || "").toString();
    let rShift = (data[i][14] || "").toString();
    let rEnteredBy = (data[i][15] || "").toString();
    if (!canViewOthers && !isAdmin && rEnteredBy !== currentUser) continue;
    if (rDate === selectedDate && data[i][13]) {
      todayTotalAll += parseFloat(data[i][13]) || 0;
    }
    if (rDate === selectedDate && rMachine === machineNo.toString() && rShift === shift) {
      if (!canViewOthers && !isAdmin && rEnteredBy !== currentUser) continue;
      targetRowIndex = i + 2;
      fullProgStr = (data[i][2] || "").toString();
      let progs = fullProgStr.split(" ... ").filter(Boolean);
      lastProgram = progs[progs.length - 1] || "";
      totalWeight = parseFloat(data[i][13]) || 0;
      enteredBy = rEnteredBy || "-";
    }
  }

  todayTotalAll = Math.round(todayTotalAll * 100) / 100;
  totalWeight = Math.round(totalWeight * 100) / 100;

  if (targetRowIndex > -1) {
    let targetRange = sheet.getRange(targetRowIndex, 4, 1, 10);
    let vals = targetRange.getValues()[0];
    let bgs = targetRange.getBackgrounds()[0];
    let lines = targetRange.getFontLines()[0];
    let colors = targetRange.getFontColors()[0];
    for (let j = 0; j < 10; j++) {
      if (vals[j] !== "" && vals[j] !== null) {
        rolls[j] = { val: vals[j], bg: bgs[j], strike: lines[j], color: colors[j] };
        nextRoll++;
      }
    }

    let progs = fullProgStr.split(" ... ").filter(Boolean);
    if (progs.length > 0) {
      const colorList = ["#ffcdd2", "#bbdefb", "#c8e6c9", "#fff9c4", "#e1bee7", "#b2ebf2", "#ffccbc", "#d1c4e9"];
      let colorMap = {};
      for (let p = 0; p < progs.length; p++) {
        colorMap[colorList[p % colorList.length]] = progs[p];
      }
      let breakdown = {};
      for (let j = 0; j < 10; j++) {
        if (rolls[j].val !== "-" && rolls[j].val !== "") {
          let bg = rolls[j].bg || "#ffffff";
          let progName = colorMap[bg] || "Unknown";
          if (!breakdown[progName]) breakdown[progName] = { weights: [], total: 0 };
          let w = parseFloat(rolls[j].val) || 0;
          if (rolls[j].color !== "#d32f2f" && rolls[j].color !== "#ff0000") {
            breakdown[progName].weights.push(w);
            breakdown[progName].total += w;
          } else {
            breakdown[progName].weights.push(w + " (R)");
          }
        }
      }
      for (let key in breakdown) {
        programBreakdown.push({
          program: key,
          weights: breakdown[key].weights,
          total: Math.round(breakdown[key].total * 100) / 100
        });
      }
    }
  }

  return {
    nextRoll: 'R' + nextRoll,
    totalWeight: totalWeight,
    todayTotal: todayTotalAll,
    lastProgram: lastProgram,
    fullProgramStr: fullProgStr,
    rollData: rolls,
    enteredBy: enteredBy,
    programBreakdown: programBreakdown
  };
}

function saveData(formObject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Production_Database');
    if (!sheet) {
      sheet = ss.insertSheet('Production_Database');
      sheet.appendRow(['Date', 'M/C NO', 'PROGRAM NO', 'R1', 'R2', 'R3', 'R4', 'R5', 'R6', 'R7', 'R8', 'R9', 'R10', 'Total Weight', 'Shift', 'Entered By']);
    }
    if (sheet.getRange(1, 15).getValue() !== "Shift") {
      sheet.getRange(1, 15).setValue("Shift");
      sheet.getRange(1, 16).setValue("Entered By");
    }
    var isAdmin = (formObject.enteredBy || "").toLowerCase() === "admin";
    var permission = formObject.permission || "ViewOnly";
    var assignedShift = formObject.assignedShift || "";
    if (!isAdmin && permission === "ViewOnly") return "You only have View permission. Cannot save data!";
    if (!isAdmin && assignedShift !== "All" && formObject.shift !== assignedShift) {
      return "You can only enter data in your assigned shift: " + assignedShift;
    }
    const dateStr = formObject.date;
    const machineNo = formObject.machineNo;
    const programNo = (formObject.programNo || "").toString().trim();
    const weight = parseFloat(formObject.weight) || 0;
    const condition = formObject.condition;
    const shift = formObject.shift;
    const enteredBy = formObject.enteredBy || "Unknown";
    const lastRow = sheet.getLastRow();
    let rowIndex = -1;
    let existingPrograms = "";
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow, 16).getValues();
      for (let i = 0; i < data.length; i++) {
        let rDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
        let rMachine = (data[i][1] || "").toString();
        let rShift = (data[i][14] || "").toString();
        if (rDate === dateStr && rMachine === machineNo.toString() && rShift === shift) {
          rowIndex = i + 2;
          existingPrograms = data[i][2] || "";
          break;
        }
      }
    }
    if (rowIndex === -1) {
      rowIndex = sheet.getLastRow() + 1;
      let dParts = dateStr.split('-');
      let safeDate = new Date(dParts[0], dParts[1] - 1, dParts[2], 10, 0, 0);
      sheet.getRange(rowIndex, 1, 1, 16).setValues([[safeDate, machineNo, "", "", "", "", "", "", "", "", "", "", "", 0, shift, enteredBy]]);
    }
    let progArray = existingPrograms ? existingPrograms.toString().split(" ... ") : [];
    if (programNo !== "" && !progArray.includes(programNo)) {
      progArray.push(programNo);
    }
    let updatedProgramStr = progArray.join(" ... ");
    let progIndex = progArray.indexOf(programNo);
    if (progIndex === -1) progIndex = 0;
    const colors = ["#ffcdd2", "#bbdefb", "#c8e6c9", "#fff9c4", "#e1bee7", "#b2ebf2", "#ffccbc", "#d1c4e9"];
    const cellColor = colors[progIndex % colors.length];
    let targetRange = sheet.getRange(rowIndex, 3, 1, 12);
    let rowData = targetRange.getValues()[0];
    let rowColors = targetRange.getFontColors()[0];
    let rowLines = targetRange.getFontLines()[0];
    let rowBgs = targetRange.getBackgrounds()[0];
    rowData[0] = updatedProgramStr;
    let emptyColOffset = -1;
    for (let k = 1; k <= 10; k++) {
      if (rowData[k] === "" || rowData[k] === null) {
        emptyColOffset = k - 1;
        break;
      }
    }
    if (emptyColOffset === -1) return "All 10 rolls are full!";
    let targetIndex = emptyColOffset + 1;
    rowData[targetIndex] = weight;
    rowBgs[targetIndex] = cellColor;
    if (condition === "Hold up") {
      rowLines[targetIndex] = "line-through";
      rowColors[targetIndex] = "#000000";
    } else if (condition === "Reject") {
      rowLines[targetIndex] = "line-through";
      rowColors[targetIndex] = "#d32f2f";
    } else {
      rowLines[targetIndex] = "none";
      rowColors[targetIndex] = "#000000";
    }
    let total = 0;
    for (let i = 1; i <= 10; i++) {
      if (rowData[i] !== "" && rowData[i] !== null && rowColors[i] !== "#d32f2f") {
        total += parseFloat(rowData[i]) || 0;
      }
    }
    total = Math.round(total * 100) / 100;
    rowData[11] = total;
    targetRange.setValues([rowData]);
    targetRange.setBackgrounds([rowBgs]);
    targetRange.setFontColors([rowColors]);
    targetRange.setFontLines([rowLines]);
    sheet.getRange(rowIndex, 15).setValue(shift);
    sheet.getRange(rowIndex, 16).setValue(enteredBy);
    return "OK|" + total;
  } catch (e) {
    return "Error: " + e.message;
  }
}

function getRollsForUpdate(date, machineNo, shift, currentUser, isAdmin, permission, assignedShift) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Database');
  if (!sheet || sheet.getLastRow() < 2) return { rolls: [], program: "" };
  if (!isAdmin && permission === "ViewOnly") return { rolls: [], program: "" };
  if (!isAdmin && assignedShift !== "All" && shift !== assignedShift) return { rolls: [], program: "" };
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow, 16).getValues();
  let rolls = [];
  let program = "";
  for (let i = 0; i < data.length; i++) {
    let rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    let rMachine = (data[i][1] || "").toString();
    let rShift = (data[i][14] || "").toString();
    let rEnteredBy = (data[i][15] || "").toString();
    if (rowDate === date && rMachine === machineNo.toString() && rShift === shift) {
      if (!isAdmin && rEnteredBy !== currentUser) return { rolls: [], program: "" };
      program = (data[i][2] || "").toString();
      let formatRange = sheet.getRange(i + 2, 4, 1, 10);
      let fontLines = formatRange.getFontLines()[0];
      let fontColors = formatRange.getFontColors()[0];
      for (let j = 0; j < 10; j++) {
        let val = data[i][j + 3];
        if (val !== "" && val !== null) {
          let cond = "OK";
          if (fontLines[j] === "line-through") {
            cond = (fontColors[j] === "#d32f2f" || fontColors[j] === "#ff0000") ? "Reject" : "Hold up";
          }
          rolls.push({ rollId: 'R' + (j + 1), weight: val, condition: cond });
        }
      }
      break;
    }
  }
  return { rolls: rolls, program: program };
}

function updateRollData(updateData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Database');
    if (!sheet) return "No sheet found!";
    let isAdmin = updateData.isAdmin || false;
    let permission = updateData.permission || "ViewOnly";
    let assignedShift = updateData.assignedShift || "";
    if (!isAdmin && permission === "ViewOnly") return "You only have View permission!";
    if (!isAdmin && assignedShift !== "All" && updateData.shift !== assignedShift) {
      return "You can only update your assigned shift: " + assignedShift;
    }
    let dateStr = updateData.date;
    let machineNo = updateData.machineNo;
    let shift = updateData.shift;
    let rollId = updateData.rollId;
    let newWeight = parseFloat(updateData.weight) || 0;
    let newCondition = updateData.condition;
    let newProgram = (updateData.programNo || "").toString().trim();
    let currentUser = updateData.currentUser || "";
    let colIndex = parseInt(rollId.replace('R', ''));
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return "No data found!";
    const data = sheet.getRange(2, 1, lastRow, 16).getValues();
    let targetRow = -1;
    let enteredBy = "";
    for (let i = 0; i < data.length; i++) {
      let rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      let rMachine = (data[i][1] || "").toString();
      let rShift = (data[i][14] || "").toString();
      if (rowDate === dateStr && rMachine === machineNo.toString() && rShift === shift) {
        targetRow = i + 2;
        enteredBy = (data[i][15] || "").toString();
        break;
      }
    }
    if (targetRow === -1) return "No data found!";
    if (!isAdmin && enteredBy !== currentUser) return "You do not have permission to update this data!";
    if (newProgram !== "") sheet.getRange(targetRow, 3).setValue(newProgram);
    let targetRange = sheet.getRange(targetRow, 4, 1, 11);
    let rowData = targetRange.getValues()[0];
    let rowColors = targetRange.getFontColors()[0];
    let rowLines = targetRange.getFontLines()[0];
    let targetIndex = colIndex - 1;
    rowData[targetIndex] = newWeight;
    if (newCondition === "Hold up") {
      rowLines[targetIndex] = "line-through";
      rowColors[targetIndex] = "#000000";
    } else if (newCondition === "Reject") {
      rowLines[targetIndex] = "line-through";
      rowColors[targetIndex] = "#d32f2f";
    } else {
      rowLines[targetIndex] = "none";
      rowColors[targetIndex] = "#000000";
    }
    let total = 0;
    for (let i = 0; i < 10; i++) {
      if (rowData[i] !== "" && rowData[i] !== null && rowColors[i] !== "#d32f2f" && rowColors[i] !== "#ff0000") {
        total += parseFloat(rowData[i]) || 0;
      }
    }
    total = Math.round(total * 100) / 100;
    rowData[10] = total;
    targetRange.setValues([rowData]);
    targetRange.setFontColors([rowColors]);
    targetRange.setFontLines([rowLines]);
    return "OK|" + total;
  } catch (e) {
    return "Error: " + e.message;
  }
}

function deleteRoll(deleteData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Database');
    if (!sheet) return "No sheet found!";
    let isAdmin = deleteData.isAdmin || false;
    let permission = deleteData.permission || "ViewOnly";
    let assignedShift = deleteData.assignedShift || "";
    if (!isAdmin && permission === "ViewOnly") return "You only have View permission!";
    if (!isAdmin && assignedShift !== "All" && deleteData.shift !== assignedShift) {
      return "You can only delete from your assigned shift: " + assignedShift;
    }
    let dateStr = deleteData.date;
    let machineNo = deleteData.machineNo;
    let shift = deleteData.shift;
    let rollId = deleteData.rollId;
    let currentUser = deleteData.currentUser || "";
    let colIndex = parseInt(rollId.replace('R', ''));
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return "No data found!";
    const data = sheet.getRange(2, 1, lastRow, 16).getValues();
    let targetRow = -1;
    let enteredBy = "";
    for (let i = 0; i < data.length; i++) {
      let rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      let rMachine = (data[i][1] || "").toString();
      let rShift = (data[i][14] || "").toString();
      if (rowDate === dateStr && rMachine === machineNo.toString() && rShift === shift) {
        targetRow = i + 2;
        enteredBy = (data[i][15] || "").toString();
        break;
      }
    }
    if (targetRow === -1) return "No data found!";
    if (!isAdmin && enteredBy !== currentUser) return "You do not have permission to delete this data!";
    let targetRange = sheet.getRange(targetRow, 4, 1, 11);
    let rowData = targetRange.getValues()[0];
    let rowColors = targetRange.getFontColors()[0];
    let rowLines = targetRange.getFontLines()[0];
    let rowBgs = targetRange.getBackgrounds()[0];
    let targetIndex = colIndex - 1;
    rowData[targetIndex] = "";
    rowColors[targetIndex] = "#000000";
    rowLines[targetIndex] = "none";
    rowBgs[targetIndex] = "#ffffff";
    let total = 0;
    for (let i = 0; i < 10; i++) {
      if (rowData[i] !== "" && rowData[i] !== null && rowColors[i] !== "#d32f2f" && rowColors[i] !== "#ff0000") {
        total += parseFloat(rowData[i]) || 0;
      }
    }
    total = Math.round(total * 100) / 100;
    rowData[10] = total;
    targetRange.setValues([rowData]);
    targetRange.setFontColors([rowColors]);
    targetRange.setFontLines([rowLines]);
    targetRange.setBackgrounds([rowBgs]);
    return "OK|" + total;
  } catch (e) {
    return "Error: " + e.message;
  }
}

function searchVlookupData(date, machineNo, shift, currentUser, canViewOthers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Database');
  let result = {
    found: false,
    rolls: Array(10).fill({val: "-", bg: "#ffffff", strike: "none", color: "#000000"}),
    total: 0,
    program: "",
    enteredBy: "-",
    shift: "-"
  };
  if (!sheet || sheet.getLastRow() < 2) return result;
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow, 16).getValues();
  let isAdmin = (currentUser || "").toLowerCase() === "admin";
  for (let i = 0; i < data.length; i++) {
    let rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    let rMachine = (data[i][1] || "").toString();
    let rShift = (data[i][14] || "").toString();
    let rEnteredBy = (data[i][15] || "").toString();
    if (rowDate === date && rMachine === machineNo.toString() && rShift === shift) {
      if (!canViewOthers && !isAdmin && rEnteredBy !== currentUser) return result;
      result.found = true;
      result.program = data[i][2];
      result.total = Math.round((parseFloat(data[i][13]) || 0) * 100) / 100;
      result.enteredBy = rEnteredBy || "-";
      result.shift = rShift;
      let formatRange = sheet.getRange(i + 2, 4, 1, 10);
      let bgs = formatRange.getBackgrounds()[0];
      let lines = formatRange.getFontLines()[0];
      let colors = formatRange.getFontColors()[0];
      for (let j = 0; j < 10; j++) {
        if (data[i][j + 3] !== "" && data[i][j + 3] !== null) {
          result.rolls[j] = { val: data[i][j + 3], bg: bgs[j], strike: lines[j], color: colors[j] };
        }
      }
      break;
    }
  }
  return result;
}

function getDailyReportData(dateStr, shiftFilter, currentUser, canViewOthers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Database');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow, 16).getValues();
  let isAdmin = (currentUser || "").toLowerCase() === "admin";
  let reportData = [];
  for (let i = 0; i < data.length; i++) {
    let rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    let rShift = (data[i][14] || "").toString();
    let rEnteredBy = (data[i][15] || "").toString();
    if (rowDate === dateStr) {
      if (shiftFilter !== "All" && rShift !== shiftFilter) continue;
      if (!canViewOthers && !isAdmin && rEnteredBy !== currentUser) continue;
      let rowObj = {
        machine: data[i][1],
        program: data[i][2] || "-",
        total: Math.round((parseFloat(data[i][13]) || 0) * 100) / 100,
        shift: rShift || "-",
        enteredBy: rEnteredBy || "-",
        rolls: Array(10).fill({val: "-", bg: "#ffffff", strike: "none", color: "#000000"})
      };
      let formatRange = sheet.getRange(i + 2, 4, 1, 10);
      let bgs = formatRange.getBackgrounds()[0];
      let lines = formatRange.getFontLines()[0];
      let colors = formatRange.getFontColors()[0];
      for (let j = 0; j < 10; j++) {
        if (data[i][j + 3] !== "" && data[i][j + 3] !== null) {
          rowObj.rolls[j] = { val: data[i][j + 3], bg: bgs[j], strike: lines[j], color: colors[j] };
        }
      }
      reportData.push(rowObj);
    }
  }
  reportData.sort((a, b) => parseInt(a.machine) - parseInt(b.machine));
  return reportData;
}
