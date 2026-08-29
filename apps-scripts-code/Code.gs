// 1. Create a custom menu when the sheet opens
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SWGOH Tools')
    .addItem('Import JSON Data', 'showImportDialog')
    .addToUi();
}

// 2. Show the HTML Dialog
function showImportDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Paste SWGOH JSON Payload');
}

// Helper to prevent Google Sheets from auto-casting strings with leading zeros to numbers
function forceText(val) {
  return (val !== undefined && val !== null) ? "'" + val : "";
}

// Helper: coerce a value (often a numeric string in the raw JSON, e.g. "19222960") to a real
// Number so quantities land in Sheets as numeric cells -- lets you SUM/sort/format them.
function toNumber(val) {
  var n = Number(val);
  return isNaN(n) ? 0 : n;
}

// Helper: look up a human-readable name from one of the GameDataMappings.gs tables,
// falling back to the raw ID (as a string) if we don't have a mapping for it yet.
// This is the "never breaks, just shows the ID" fallback -- new gear/materials/currencies
// added to the game after this mapping snapshot was generated will just show their raw ID
// until GameDataMappings.gs is regenerated.
function lookupName(table, id) {
  if (id === undefined || id === null || id === "") return "";
  var name = table[id];
  return name !== undefined ? name : String(id);
}

// Helper: format a raw epoch-milliseconds string (as found in expireTime) as a readable date.
function formatEpochMillis(msString) {
  var ms = Number(msString);
  if (!ms) return "";
  return Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
}

// 3. Process the JSON and write to sheets
function processSWGOHData(jsonString) {
  try {
    var data = JSON.parse(jsonString);
    var inventory = data.inventory || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Process Gear
    var gearData = [['Gear ID', 'Name', 'Quantity']];
    if (inventory.equipment) {
      inventory.equipment.forEach(function(item) {
        gearData.push([
          forceText(item.id),
          lookupName(EQUIPMENT_NAMES, item.id),
          toNumber(item.quantity)
        ]);
      });
    }
    writeToSheet(ss, 'Gear', gearData);

    // Process Currencies
    var currencyData = [['Currency ID', 'Name', 'Quantity']];
    if (inventory.currencyItem) {
      inventory.currencyItem.forEach(function(item) {
        currencyData.push([
          forceText(item.currency),
          lookupName(CURRENCY_NAMES, item.currency),
          toNumber(item.quantity)
        ]);
      });
    }
    writeToSheet(ss, 'Currency', currencyData);

    // Process Materials
    var matData = [['Material ID', 'Name', 'Category', 'Quantity']];
    if (inventory.material) {
      inventory.material.forEach(function(item) {
        var mat_id = item.id;
        var category = "General Material";
        if (mat_id.indexOf("unitshard_") === 0) category = "Unit Shard";
        else if (mat_id.indexOf("SCV_") === 0 || mat_id.indexOf("era_upgrade_") === 0 || mat_id.indexOf("RM_") === 0 || mat_id === "coaxium") category = "Relic Material";
        else if (mat_id.indexOf("datacron_") === 0) category = "Datacron Material";
        else if (mat_id.indexOf("MOD_SLICING_") === 0) category = "Mod Material";

        matData.push([
          forceText(mat_id),
          lookupName(MATERIAL_NAMES, mat_id),
          category,
          toNumber(item.quantity)
        ]);
      });
    }
    writeToSheet(ss, 'Materials', matData);

    // Process Mods
    var modData = [[
      'Mod ID', 'Definition ID', 'Level', 'Tier', 'Locked', 'Rerolled Count',
      'Primary Stat ID', 'Primary Stat Name', 'Primary Stat Value',
      'Sec 1 ID', 'Sec 1 Name', 'Sec 1 Value', 'Sec 1 Rolls',
      'Sec 2 ID', 'Sec 2 Name', 'Sec 2 Value', 'Sec 2 Rolls',
      'Sec 3 ID', 'Sec 3 Name', 'Sec 3 Value', 'Sec 3 Rolls',
      'Sec 4 ID', 'Sec 4 Name', 'Sec 4 Value', 'Sec 4 Rolls'
    ]];

    if (inventory.unequippedMod) {
      inventory.unequippedMod.forEach(function(mod) {
        var pStat = (mod.primaryStat && mod.primaryStat.stat) ? mod.primaryStat.stat : {};
        var row = [
          forceText(mod.id), forceText(mod.definitionId), mod.level, mod.tier, mod.locked, mod.rerolledCount || 0,
          pStat.unitStatId || "", lookupName(MOD_STAT_NAMES, pStat.unitStatId), pStat.unscaledDecimalValue / 100000000 || ""
        ];

        var secs = mod.secondaryStat || [];
        for (var i = 0; i < 4; i++) {
          if (i < secs.length) {
            var sStat = secs[i].stat || {};
            row.push(
              sStat.unitStatId || "",
              lookupName(MOD_STAT_NAMES, sStat.unitStatId),
              sStat.unscaledDecimalValue / 100000000 || "",
              secs[i].statRolls || 0
            );
          } else {
            row.push("", "", "", "");
          }
        }
        modData.push(row);
      });
    }
    writeToSheet(ss, 'Mods', modData);

    // Process Lightspeed Tokens
    var lstData = [['Token Instance ID', 'Type', 'Name', 'Target', 'Stars', 'Relic Level', 'Gear Tier', 'Character Level', 'Expires']];
    if (inventory.lightspeedToken) {
      inventory.lightspeedToken.forEach(function(token) {
        var def = LST_DEFINITIONS[token.definitionId];
        lstData.push([
          forceText(token.id),
          forceText(token.definitionId),
          def ? def.name : String(token.definitionId),
          def ? def.target : "",
          def ? def.stars : "",
          def ? def.relicLevel : "",
          def ? def.gearTier : "",
          def ? def.level : "",
          formatEpochMillis(token.expireTime)
        ]);
      });
    }
    writeToSheet(ss, 'Lightspeed Tokens', lstData);

    return "Success! Wrote to " + (gearData.length - 1) + " gear items, " + (matData.length - 1) + " materials, " + (modData.length - 1) + " mods, and " + (lstData.length - 1) + " lightspeed tokens.";

  } catch (e) {
    return "Error parsing JSON: " + e.toString();
  }
}

// Helper to create or clear a sheet, then write a 2D array to it.
// Always writes (even when dataArray is just the header row) so a category that's gone
// to zero -- e.g. you use your last Lightspeed Token -- correctly clears out the sheet
// instead of leaving stale rows from a previous import sitting there.
function writeToSheet(spreadsheet, sheetName, dataArray) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
}