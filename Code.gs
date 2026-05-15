/**
 * Serves the HTML for the web application.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Construction Cost App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Fetches initial data for the app from Google Sheets.
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bidsSheet = ss.getSheetByName('Bids');
  const equipmentSheet = ss.getSheetByName('Equipment');
  const personnelSheet = ss.getSheetByName('Personnel');

  const projects = bidsSheet ? bidsSheet.getDataRange().getValues().slice(1) : [];
  const equipment = equipmentSheet ? equipmentSheet.getDataRange().getValues().slice(1).map(row => row[0]) : [];
  const personnel = personnelSheet ? personnelSheet.getDataRange().getValues().slice(1).map(row => ({
    name: row[0],
    rate: row[1]
  })) : [];

  return {
    projects: projects.map(row => ({ name: row[0], bid: row[1] })),
    equipment: equipment,
    personnel: personnel
  };
}

/**
 * Calculates labor cost based on California rules and labor tax.
 */
function calculateLaborCost(hours, rate, isSickDay) {
  let cost = 0;
  if (isSickDay) {
    cost = 8 * rate;
  } else {
    if (hours <= 8) {
      cost = hours * rate;
    } else if (hours <= 12) {
      cost = (8 * rate) + ((hours - 8) * rate * 1.5);
    } else {
      cost = (8 * rate) + (4 * rate * 1.5) + ((hours - 12) * rate * 2);
    }
  }

  // Add 32% labor tax
  return cost * 1.32;
}

/**
 * Handles the submission of the 4-step data entry flow.
 */
function submitData(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('DailyLogs');
  if (!logSheet) {
    logSheet = ss.insertSheet('DailyLogs');
    logSheet.appendRow(['Timestamp', 'Project', 'Equipment', 'Labor Total', 'Materials Total', 'Daily Total']);
  }

  const { project, equipment, labor, materials } = payload;

  // Calculate Labor Total
  let laborTotal = 0;
  labor.forEach(entry => {
    const hours = entry.isSickDay ? 8 : (new Date(entry.endTime) - new Date(entry.startTime)) / (1000 * 60 * 60);
    laborTotal += calculateLaborCost(hours, entry.rate, entry.isSickDay);
  });

  // Calculate Materials Total
  let materialsTotal = 0;
  materials.forEach(item => {
    materialsTotal += item.quantity * item.price;
  });

  const dailyTotal = laborTotal + materialsTotal;

  // Save to sheet
  logSheet.appendRow([
    new Date(),
    project.name,
    equipment.join(', '),
    laborTotal,
    materialsTotal,
    dailyTotal
  ]);

  // Budget Logic: Red (>90% bid), Yellow (50-90%), Green (<50%)
  // Note: For simplicity, we compare daily total vs total bid.
  // In a real app, you'd sum all logs for this project.
  const bidAmount = project.bid;
  const percentage = (dailyTotal / bidAmount) * 100;

  let status = 'Green';
  if (percentage > 90) {
    status = 'Red';
  } else if (percentage >= 50) {
    status = 'Yellow';
  }

  return {
    status: status,
    percentage: percentage.toFixed(2),
    dailyTotal: dailyTotal.toFixed(2)
  };
}
