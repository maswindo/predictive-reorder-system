// Predictive Reorder System
// Calculates projected usage and stock needed for upcoming months

function updateProjectedUsageAndReconciliation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var auditLogSheet = ss.getSheetByName("Transfer Audit Log");
  var projectedUsageSheet = ss.getSheetByName("Projected Inventory Usage");
  var locationInventorySheet = ss.getSheetByName("Location-Based Inventory");

  if (!auditLogSheet || !projectedUsageSheet || !locationInventorySheet) {
    Logger.log("Error: One or more required sheets not found.");
    return;
  }

  var auditData = auditLogSheet.getDataRange().getValues();
  var projectedUsageData = projectedUsageSheet.getDataRange().getValues();
  var locationInventoryData = locationInventorySheet.getDataRange().getValues();

  var currentDate = new Date();
  var sixMonthsAgo = new Date();
  sixMonthsAgo.setMonth(currentDate.getMonth() - 6);

  var materialUsage = {};
  var last3MonthsUsage = {};
  var previous3MonthsUsage = {};
  var warehouseStockLevels = {};

  for (var i = 1; i < locationInventoryData.length; i++) {
    var locationName = String(locationInventoryData[i][0]).trim();
    var materialCode = String(locationInventoryData[i][1]).trim();
    var stock = parseFloat(locationInventoryData[i][2]) || 0;

    if (locationName.startsWith("Meerkat")) {
      if (!warehouseStockLevels[materialCode]) warehouseStockLevels[materialCode] = 0;
      warehouseStockLevels[materialCode] += stock;
    }
  }

  for (var i = 1; i < auditData.length; i++) {
    var transferDate = new Date(auditData[i][0]);
    var materialCode = String(auditData[i][1]).trim();
    var quantity = parseFloat(auditData[i][8]);

    if (!materialCode || isNaN(quantity)) continue;

    if (!materialUsage[materialCode]) materialUsage[materialCode] = 0;

    if (transferDate >= sixMonthsAgo) materialUsage[materialCode] += quantity;

    var last3Months = new Date();
    last3Months.setMonth(currentDate.getMonth() - 3);
    var previous3Months = new Date();
    previous3Months.setMonth(currentDate.getMonth() - 6);

    if (transferDate >= last3Months) {
      if (!last3MonthsUsage[materialCode]) last3MonthsUsage[materialCode] = 0;
      last3MonthsUsage[materialCode] += quantity;
    } else if (transferDate >= previous3Months) {
      if (!previous3MonthsUsage[materialCode]) previous3MonthsUsage[materialCode] = 0;
      previous3MonthsUsage[materialCode] += quantity;
    }
  }

  for (var i = 1; i < projectedUsageData.length; i++) {
    var materialCode = String(projectedUsageData[i][0]).trim();
    var totalUsage = materialUsage[materialCode] || 0;
    var avgMonthlyUsage = totalUsage / 6;

    projectedUsageSheet.getRange(i + 1, 3).setValue(totalUsage);
    projectedUsageSheet.getRange(i + 1, 4).setValue(avgMonthlyUsage);

    var projectedUsage = avgMonthlyUsage;
    if (last3MonthsUsage[materialCode] && previous3MonthsUsage[materialCode] && previous3MonthsUsage[materialCode] > 0) {
      var trendFactor = (last3MonthsUsage[materialCode] - previous3MonthsUsage[materialCode]) / previous3MonthsUsage[materialCode];
      trendFactor = Math.min(trendFactor, 1.0);
      projectedUsage = avgMonthlyUsage * (1 + trendFactor);
    }

    if (totalUsage === 0) projectedUsage = 0;
    projectedUsageSheet.getRange(i + 1, 5).setValue(projectedUsage);

    var percentChange = "N/A";
    if (previous3MonthsUsage[materialCode] > 0) {
      percentChange = (last3MonthsUsage[materialCode] - previous3MonthsUsage[materialCode]) / previous3MonthsUsage[materialCode];
    }
    projectedUsageSheet.getRange(i + 1, 6).setValue(percentChange);

    var stockOnHand = warehouseStockLevels[materialCode] || 0;
    projectedUsageSheet.getRange(i + 1, 7).setValue(stockOnHand);

    var stockNeeded = projectedUsage > stockOnHand ? projectedUsage - stockOnHand : 0;
    projectedUsageSheet.getRange(i + 1, 8).setValue(stockNeeded);
  }

  Logger.log("Projected Inventory Usage update completed.");
}
