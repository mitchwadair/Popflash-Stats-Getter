function sortBySelectedCol() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getCurrentCell();
  var cellRange = sheet.getRange("O1:Z1");
  
  if (cell.getRow() == 1) {
    var rangeCol1 = cellRange.getColumn();
    var rangeCol2 = cellRange.getLastColumn();
    var col = cell.getColumn();
    if (col > rangeCol1 && col < rangeCol2) {
      var lastRow = sheet.getLastRow();
      var leaderboardRange = sheet.getRange("O2:Z" + lastRow);
      leaderboardRange.sort({column:col, ascending:false});
    }
  }
}

function getStats() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getRange("M2:M").getValues();
  var dataLength = data.filter(String).length;
  var links = sheet.getRange("M2:M" + (dataLength+1)).getValues();
  var stats = [];
  links.forEach(function(row) {
    var newStats = getStatsForMatch(row[0]);
    stats = mergeStats(stats, newStats);
  });
  
  //clear the leaderboard
  var lastRow = sheet.getLastRow();
  var leaderboardRange = sheet.getRange("O1:Z" + lastRow);
  leaderboardRange.clear();
  
  var row = sheet.getRange("O1:Z1");
  row.setFontWeight("bold");
  row.setBackgroundRGB(200, 200, 200);
  row.setValues([['Name ', 'Rating ', 'Kills ', 'Assists ', 'Deaths ', 'ADR ', 'Headshot % ', 'Clutch Kills ', 'Bombs Planted ', 'Bombs Defused ', 'Flash Assists ', 'Enemy Flash Duration ']])
  
  for (var i = 0; i < stats.length; i++) {
    row = sheet.getRange("O" + (i+2) + ":Z" + (i+2));
    row.setBackgroundRGB(235, 235, 235);
    row.setValues([[stats[i].name, stats[i].rating, stats[i].kills, stats[i].assists, stats[i].deaths, stats[i].adr, stats[i].hsp, stats[i].ck, stats[i].bp, stats[i].bd, stats[i].fa, stats[i].fed]])
  }
  
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  
  lastRow = sheet.getLastRow();
  leaderboardRange = sheet.getRange("O2:Z" + lastRow);
  var colToSortBy = sheet.getRange("A1:P1").getLastColumn();
  leaderboardRange.sort({column:colToSortBy, ascending:false});
}

//get stats for entire match page given a url
function getStatsForMatch(url) {
  var pageContent = UrlFetchApp.fetch(url).getContentText();
  var fixedContent = pageContent.replace(new RegExp("<br>", 'g'), "<br></br>");
  var parsedContent = XmlService.parse(fixedContent.substring(fixedContent.indexOf("<body"), fixedContent.indexOf("</body") + 7));
  var html = parsedContent.getRootElement();
  var scoreboards = getElementsByTag(getElementsByClassName(html, 'scoreboards')[0], 'table');
  var stats = [];
  for (var i = 0; i < scoreboards.length; i++) {
    var rows = getElementsByTag(scoreboards[i], 'tr');
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      var statsObject = buildStatsObjectForRow(row);
      stats.push(statsObject);
    }
  }
  return stats;
}

//create a JSON object with stats for a specific row of data
function buildStatsObjectForRow(r) {
  var data = getElementsByTag(r, 'td');
  var categories = ['name', 'id', 'kills', 'assists', 'deaths', 'fa', 'adr', 'rating', 'hsp', 'ck', 'bp', 'bd', 'fed'];
  var values = [];
  for (var i = 0; i < data.length; i++) {
    var entry = data[i];
    if (i == 0) {
      values.push(entry.getAttribute('title').getValue())
      values.push(getElementsByTag(entry, 'a')[0].getAttribute('href').getValue());
    } else {
      values.push(entry.getValue());
    }
  }
  var obj = {};
  for (var i = 0; i < categories.length; i++) {
    obj[categories[i]] = values[i];
  }
  return obj;
}

function matchIds(player) {
  return player.id == this.id;
}

//merge two arrays of stats objects
function mergeStats(oldStats, newStats) {
  var mergedStats = [];
  if (oldStats.length > 0) {
    for (var i = 0; i < oldStats.length; i++) {
      var match = newStats.filter(matchIds, oldStats[i]);
      if (match.length == 0) {
        mergedStats.push(oldStats[i]);
      } else {
        mergedStats.push(mergePlayerStats(oldStats[i], match[0]));
      }
    }
    for (var i = 0; i < newStats.length; i++) {
      var match = oldStats.filter(matchIds, newStats[i]);
      if (match.length == 0) {
        mergedStats.push(newStats[i]);
      }
    }
  } else {
    mergedStats = newStats;
  }
  return mergedStats;
}

//take an object and merge the stats, adding and averaging where needed
function mergePlayerStats(ops, nps) {
  var obj = {};
  obj.name = nps.name;
  obj.id = nps.id;
  obj.kills = Number(ops.kills) + Number(nps.kills);
  obj.assists = Number(ops.assists) + Number(nps.assists);
  obj.deaths = Number(ops.deaths) + Number(nps.deaths);
  obj.fa = Number(ops.fa) + Number(nps.fa);
  obj.adr = ((Number(ops.adr) + Number(nps.adr))/2).toFixed(2);
  obj.rating = ((Number(ops.rating) + Number(nps.rating))/2).toFixed(2);
  obj.hsp = ((Number(ops.hsp) + Number(nps.hsp))/2).toFixed(2);
  obj.ck = Number(ops.ck) + Number(nps.ck);
  obj.bp = Number(ops.bp) + Number(nps.bp);
  obj.bd = Number(ops.bd) + Number(nps.bd);
  obj.fed = (Number(ops.fed) + Number(nps.fed)).toFixed(2);
  return obj;
}

//useful functions thanks to Google:
//https://sites.google.com/site/scriptsexamples/learn-by-example/parsing-html
function getElementsByClassName(element, classToFind) {
  var data = [];
  var descendants = element.getDescendants();
  descendants.push(element);  
  for(i in descendants) {
    var elt = descendants[i].asElement();
    if(elt != null) {
      var classes = elt.getAttribute('class');
      if(classes != null) {
        classes = classes.getValue();
        if(classes == classToFind) data.push(elt);
        else {
          classes = classes.split(' ');
          for(j in classes) {
            if(classes[j] == classToFind) {
              data.push(elt);
              break;
            }
          }
        }
      }
    }
  }
  return data;
}
function getElementsByTag(element, tagName) {  
  var data = [];
  var descendants = element.getDescendants();  
  for(i in descendants) {
    var elt = descendants[i].asElement();     
    if( elt !=null && elt.getName()== tagName) data.push(elt);      
  }
  return data;
}
