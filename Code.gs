/*
	Sort the leaderboard by the selected category column
	This will only work if the selected cell is within the range of cells that contain leaderboard categories
*/
function sortBySelectedCol() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getCurrentCell(); //get the active cell (i.e. the category to sort by)
  var cellRange = sheet.getRange("O1:AA1"); //get the range of category cells
  
  //ensure the selected cell is in the first row (the row which contains categories)
  if (cell.getRow() == 1) {
    var rangeCol1 = cellRange.getColumn(); //get the first column of the category range
    var rangeCol2 = cellRange.getLastColumn(); //get the last column of the category range
    var col = cell.getColumn(); //get the column of the current selected cell
	//ensure the selected cell is within the category range
    if (col >= rangeCol1 && col <= rangeCol2) {
	  //get the range of the leaderboard
      var lastRow = sheet.getLastRow();
      var leaderboardRange = sheet.getRange("O2:AA" + lastRow);
	  
	  //sort the leaderboard by the selected category in descending order (highest to lowest)
      leaderboardRange.sort({column:col, ascending:false});
    }
  }
}

/*
	Fetch the stats from popflash, wrangle the data, then populate the leaderboard
*/
function getStats() {
  var sheet = SpreadsheetApp.getActiveSheet();
  //get popflash links
  var data = sheet.getRange("M2:M").getValues(); //M is the column which contains popflash links.
  var dataLength = data.filter(String).length; //get the number of links in the column
  var links = sheet.getRange("M2:M" + (dataLength+1)).getValues(); //get an array of links
  
  var stats = []; //declare an empty array to store our stats
  //for every link, get the stats for the match, and merge them with existing stats
  links.forEach(function(row) {
    var newStats = getStatsForMatch(row[0]); //get the stats
    stats = mergeStats(stats, newStats); //merge them with stats from previous iterations
  });
  
  //clear the leaderboard
  var lastRow = sheet.getLastRow(); //get the last row of the sheet, so we can get our range
  var leaderboardRange = sheet.getRange("O1:AA" + lastRow); //get the range of the leaderboard (default placement is cols O to AA)
  leaderboardRange.clear(); //clear the leaderboard range to populate fresh
  
  //begin creating our new leaderboard
  var row = sheet.getRange("O1:AA1"); //the category row
  row.setFontWeight("bold");
  row.setBackgroundRGB(200, 200, 200);
  //set the values in the category row using this array of values
  row.setValues([['Name ', 'Games Played ', 'Rating ', 'Kills ', 'Assists ', 'Deaths ', 'ADR ', 'Headshot % ', 'Clutch Kills ', 'Bombs Planted ', 'Bombs Defused ', 'Flash Assists ', 'Enemy Flash Duration ']])
  
  //for each value in our array of stats, populate another row of data
  for (var i = 0; i < stats.length; i++) {
    row = sheet.getRange("O" + (i+2) + ":AA" + (i+2)); //get the corresponding row range to populate
    row.setBackgroundRGB(235, 235, 235);
	//set the values of the row in order as above, using data from the stats array
    row.setValues([[stats[i].name, stats[i].gp, stats[i].rating, stats[i].kills, stats[i].assists, stats[i].deaths, stats[i].adr, stats[i].hsp, stats[i].ck, stats[i].bp, stats[i].bd, stats[i].fa, stats[i].fed]])
  }
  
  sheet.autoResizeColumns(1, sheet.getLastColumn()); //resize all columns to ensure readability
  
  //get the range of our new leaderboard
  lastRow = sheet.getLastRow();
  leaderboardRange = sheet.getRange("O2:AA" + lastRow);
  
  //sort the leaderboard by best to worst Rating
  var colToSortBy = sheet.getRange("A1:Q1").getLastColumn(); //the P column is Rating by default
  leaderboardRange.sort({column:colToSortBy, ascending:false}); //sort the leaderboard
}

//get stats for entire match page given a url
function getStatsForMatch(url) {
  //get the page contents
  var pageContent = UrlFetchApp.fetch(url).getContentText(); //fetch the popflash page content
  var fixedContent = pageContent.replace(new RegExp("<br>", 'g'), "<br></br>"); //replace any instances where we have a "<br>" tag with a closed variant (prevents XmlService error)
  var parsedContent = XmlService.parse(fixedContent.substring(fixedContent.indexOf("<body"), fixedContent.indexOf("</body") + 7)); //parse the body of the page contents
  var html = parsedContent.getRootElement(); //get the root element of the parsed content
  var scoreboards = getElementsByTag(getElementsByClassName(html, 'scoreboards')[0], 'table'); //get the scoreboards
  
  var stats = []; //declare our empty array of stats
  //for each scoreboard, get the data and add it to our stats array
  for (var i = 0; i < scoreboards.length; i++) {
    var rows = getElementsByTag(scoreboards[i], 'tr'); //get the rows of data
	//for each row of data, create a new stats object and add it to our stats array
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      var statsObject = buildStatsObjectForRow(row); //build our stats object from the row data
      stats.push(statsObject); //add the object to our array
    }
  }
  return stats; //return our array of stats
}

//create a JSON object with stats for a specific row of data
function buildStatsObjectForRow(r) {
  var data = getElementsByTag(r, 'td'); //get the cells of data from the row passed into the function
  var categories = ['name', 'id', 'kills', 'assists', 'deaths', 'fa', 'adr', 'rating', 'hsp', 'ck', 'bp', 'bd', 'fed']; //an array of categories to map to
  var values = []; //declare our empty array of values
  //for each data entry, add it to our array of values
  for (var i = 0; i < data.length; i++) {
    var entry = data[i]; //get our data entry
	//if this is the first run through our loop, our data case if slightly different
    if (i == 0) {
      values.push(entry.getAttribute('title').getValue()) //get the player's name from the title attribute of the cell
      values.push(getElementsByTag(entry, 'a')[0].getAttribute('href').getValue()); //get the popflash user id from the cell, this is stored in an "<a>" tag to link to a profile
    } else {
      values.push(entry.getValue()); //add the data entry to the array of values
    }
  }
  
  //build our object
  var obj = {}; //declare our empty object
  //for every category, add a new key to the object with its corresponding value
  for (var i = 0; i < categories.length; i++) {
    obj[categories[i]] = values[i]; //create a key from the category, and assign its value based on the corresponding index in values
  }
  obj.gp = 1; //new stats object, so initialize games played to 1
  
  return obj; //return our stats object
}

/*
	function to call from Arrays.filter
	returns true if the player ids match
*/
function matchIds(player) {
  return player.id == this.id;
}

//merge two arrays of stats objects
function mergeStats(oldStats, newStats) {
  var mergedStats = []; //declare an empty array to add merged stats to
  //if there there are stats to merge, do it.  Otherwise we just add the new stats to our stats array
  if (oldStats.length > 0) {
    mergedStats = oldStats; //begin mergedStats as the old array of stats
	//while we still have entries in our new stats, compare and merge
    while (newStats.length > 0) {
      var compare = newStats.pop(); //get a stats object to compare
      var match = mergedStats.filter(matchIds, compare); //find a matching player in mergedStats
	  //if we have a match, merge and replace.  If not, add to the array
      if (match.length > 0) {
        var index = mergedStats.indexOf(match[0]); //get the index of the matching stats object
        mergedStats[index] = mergePlayerStats(mergedStats[index], compare); //replace the object with the merged stats
      } else {
        mergedStats.push(compare); //push the unmatched stats object to the merged array
      }
    }
  } else {
    mergedStats = newStats; //first run through, set merged to the new stats
  }
  
  return mergedStats; //return our merged stats array
}

//take an object and merge the stats, adding and averaging where needed
function mergePlayerStats(ops, nps) {
  var obj = {};
  obj.name = nps.name; //take the newest name for the player
  obj.id = nps.id; //set the player id (used to match players)
  obj.kills = Number(ops.kills) + Number(nps.kills); //add kills together
  obj.assists = Number(ops.assists) + Number(nps.assists); //add assists together
  obj.deaths = Number(ops.deaths) + Number(nps.deaths); //add deaths together
  obj.fa = Number(ops.fa) + Number(nps.fa); //add flash assists together
  obj.adr = ((Number(ops.adr) + Number(nps.adr))/2).toFixed(2); //find the average ADR and round to 2 decimal places
  obj.rating = ((Number(ops.rating) + Number(nps.rating))/2).toFixed(2); //find the average player Rating and round to 2 decimal places
  obj.hsp = ((Number(ops.hsp) + Number(nps.hsp))/2).toFixed(2); //find the average headshot percentage and round to 2 decimal places
  obj.ck = Number(ops.ck) + Number(nps.ck); //add cluch kills together
  obj.bp = Number(ops.bp) + Number(nps.bp); //add bomb plants together
  obj.bd = Number(ops.bd) + Number(nps.bd); //add bomb defuses together
  obj.fed = (Number(ops.fed) + Number(nps.fed)).toFixed(2); //add flash enemy duration together, round to w decimal places
  obj.gp = ops.gp + 1; //add one to games played
  return obj; //return our merged stats
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
