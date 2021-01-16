function onOpen(e) {

  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("NexusHub Price Check")
      .addItem("Start", 'startMenu')
      .addToUi();
}

function startMenu()
{
  var html = HtmlService.createTemplateFromFile("FactionSidebarHtml")
    .evaluate()
    .setTitle("Price check")
    .setWidth(400);

  SpreadsheetApp.getUi().showSidebar(html);
}

function fillPriceData(region, server)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName("Data");
  var data = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues();
  var props = PropertiesService.getDocumentProperties();
  
  props.setProperties({
    "region": region,
    "server": server
  });
  
  var objData = getObjectFromArray(data); 
 
  for(var i = 0; i < objData.length; i++)
  {
    var allyData = getItemData(server, "alliance", objData[i].itemId);
    var hordeData = getItemData(server, "horde", objData[i].itemId);
    var currentRow = i + 2;
    
    sheet.getRange(currentRow, 2).setValue(allyData.name);
    
    if(isDataAvailable(allyData.stats.current))
    {
      sheet.getRange(currentRow, 3).setValue(allyData.stats.current.marketValue / 10000);
      sheet.getRange(currentRow, 4).setValue(allyData.stats.current.minBuyout / 10000);
    }
    else
    {
      sheet.getRange(currentRow, 3).setValue("No data to show");
      sheet.getRange(currentRow, 4).setValue("No data to show");
    }
    
    if(isDataAvailable(hordeData.stats.current))
    {
      sheet.getRange(currentRow, 5).setValue(hordeData.stats.current.marketValue / 10000);
      sheet.getRange(currentRow, 6).setValue(hordeData.stats.current.minBuyout / 10000);
    }
    else
    {
      sheet.getRange(currentRow, 5).setValue("No data to show");
      sheet.getRange(currentRow, 6).setValue("No data to show");
    }
    
    sheet.getRange(currentRow, 7).setValue(allyData.stats.lastUpdated);
    sheet.getRange(currentRow, 13).setFormula("=HYPERLINK(CONCAT(\"https://nexushub.co/wow-classic/items/" + server + "-horde/\" ; A" +  currentRow + "), \"Link\")");
    sheet.getRange(currentRow, 14).setFormula("=HYPERLINK(CONCAT(\"https://nexushub.co/wow-classic/items/" + server + "-alliance/\" ; A" +  currentRow + "), \"Link\")");
    
    Utilities.sleep(250);
  }
  
  var nonHeaderRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  nonHeaderRange.sort({column: 11, ascending: false});
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

function getProperties()
{
  var props = PropertiesService.getDocumentProperties();
  
  var data = {
    "region": props.getProperty("region") || "US",
    "server": props.getProperty("server") || "anathema",
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
