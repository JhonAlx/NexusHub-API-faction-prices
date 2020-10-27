function onOpen(e) {

  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("NexusHub Price Check")
      .addItem("Start", 'startMenu')
      .addToUi();
}

function startMenu()
{
  var html = HtmlService.createTemplateFromFile("ServerSidebarHtml")
    .evaluate()
    .setTitle("Price check")
    .setWidth(400);

  SpreadsheetApp.getUi().showSidebar(html);
}

function fillPriceData(region, server1, faction1, server2, faction2)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName("Data");
  var data = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues();
  var props = PropertiesService.getDocumentProperties();
  
  props.setProperties({
    "region": region,
    "factionOne": faction1,
    "factionTwo": faction2,
    "serverOne": server1,
    "serverTwo": server2
  });
  
  sheet.getRange("C1").setValue(capitalize(server1) + "-" + capitalize(faction1) + " MV");
  sheet.getRange("D1").setValue(capitalize(server1) + "-" + capitalize(faction1)+ " min buyout");
  
  sheet.getRange("E1").setValue(capitalize(server2) + "-" + capitalize(faction2) + " MV");
  sheet.getRange("F1").setValue(capitalize(server2) + "-" + capitalize(faction2) + " min buyout");
  
  var objData = getObjectFromArray(data); 
 
  for(var i = 0; i < objData.length; i++)
  {
    var server1Data = getItemData(server1, faction1, objData[i].itemId);
    var server2Data = getItemData(server2, faction2, objData[i].itemId);
    var currentRow = i + 2;
    
    sheet.getRange(i + 2, 2).setValue(server1Data.name);
    
    if(isDataAvailable(server1Data.stats.current))
    {
      sheet.getRange(i + 2, 3).setValue(server1Data.stats.current.marketValue / 10000);
      sheet.getRange(i + 2, 4).setValue(server1Data.stats.current.minBuyout / 10000);
    }
    else
    {
      sheet.getRange(i + 2, 3).setValue("No data to show");
      sheet.getRange(i + 2, 4).setValue("No data to show");
    }
    
    if(isDataAvailable(server2Data.stats.current))
    {
      sheet.getRange(i + 2, 5).setValue(server2Data.stats.current.marketValue / 10000);
      sheet.getRange(i + 2, 6).setValue(server2Data.stats.current.minBuyout / 10000);
    }
    else
    {
      sheet.getRange(i + 2, 5).setValue("No data to show");
      sheet.getRange(i + 2, 6).setValue("No data to show");
    }
    
    sheet.getRange(i + 2, 7).setValue(server1Data.stats.lastUpdated);
    sheet.getRange(i + 2, 8).setFormula("=IF(F" + currentRow + "<D" + currentRow + ", \"" + capitalize(server2) + "\", \"" + capitalize(server1) + "\")")
    sheet.getRange(i + 2, 9).setFormula("=IF(H" + currentRow + "=\"" + capitalize(server1) + "\",F" + currentRow + "-D" + currentRow + ",D" + currentRow + "-F" + currentRow + ")")
  }
  
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function isDataAvailable(itemData)
{
  return itemData != null;
}

function getObjectFromArray(data)
{
  let objData = [];

  for (var i = 0; i < data.length - 1; i++) {
    var itemId = data[i][0].toString();
    var entry = {
      "itemId": itemId,
    }

    objData.push(entry);
  }

  return objData;
}

function capitalize(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}

function getProperties()
{
  var props = PropertiesService.getDocumentProperties();
  
  var data = {
    "region": props.getProperty("region") || "US",
    "factionOne": props.getProperty("factionOne") || "horde",
    "factionTwo": props.getProperty("factionTwo") || "horde",
    "serverOne": props.getProperty("serverOne") || "anathema",
    "serverTwo": props.getProperty("serverTwo") || "arcanite-reaper",
  };
  
  return data;
}

function getItemData(server, faction, itemId) 
{
  var query = "/wow-classic/v1/items/" + server + "-" + faction + "/" + itemId;
  var url = "https://api.nexushub.co" + query;
  
  var response = UrlFetchApp.fetch(url);
  var itemObject = JSON.parse(response);
  
  return itemObject;
}
