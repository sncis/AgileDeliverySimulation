//The onOpen function is executed automatically every time a Spreadsheet is loaded
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //define new menu as an empty list
  var menuEntries = [];

  //creating new menu point
  menuEntries.push({name: 'Run Simulation', functionName: 'monteCarlo',});

  //adding the menu point to spreadsheet menu
  ss.addMenu('Run', menuEntries);

  //creating new sheets for writing data insight
  function createSheet(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var sheetname =["Burn-up","Simulations","Scratch"];

    //set name for the first sheet
    sheets[0].setName("Data");

    //creating new sheets
    for(var i = 1; i <= sheetname.length; i ++){
      //check if sheets already exist
      if(sheets[i] == null){
        //inserting new sheets
        ss.insertSheet(sheetname[i - 1], i);
        Logger.log(sheetname[i-1]);
      }

    }
  }
  createSheet();
}

// function to calculating the average of how many weeks n stories takes
function monteCarlo(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var simsSheet = ss.getSheets()[2];

  //deleting all existing charts
  var allCharts = simsSheet.getCharts();
  for (var i in allCharts) {
    simsSheet.removeChart(allCharts[i]);
  }
  simsSheet.clear();

  // Todo: parameterize range inputs
  //defining the range in datasheet which we want to consider in the calculation
  var storiesRange = dataSheet.getRange("F3:F34");
  var weeksRange = dataSheet.getRange("G3:G34");

  //getting the values out of the ranges
  var stories = storiesRange.getValues();
  var weeks = weeksRange.getValues();

  // input paraeter which can be defined by user
  var targetStories = dataSheet.getRange("K12").getValue();
  var fractionOfTeamCapacity = dataSheet.getRange("K13").getValue();

  var deliveredStories = 0;
  var elapsedWeeks = 0;

  var sims = [];
  var simsCount = 0;
  var simsTotal = 10000; // simulation iterations

  function roundToTenths(n) {
    n = n * 10;
    n = Math.floor(n);
    n = n / 10;
    return n;
  }
  //calcualting the median
  function median(values){
    values.sort(function(a,b){
    return a-b;
    });

    if(values.length ===0) return 0

    var half = Math.floor(values.length / 2);

    if (values.length % 2)
      return values[half];
    else
      return (values[half - 1] + values[half]) / 2.0;
  }
  //calculating the elapsed weeks
  while (simsCount < simsTotal) {
    deliveredStories = 0; //reset the number to 0 after every iteration
    elapsedWeeks = 0;
    while (deliveredStories < targetStories) {
      var idx = Math.floor(Math.random() * stories.length);
      deliveredStories += stories[idx][0];
      elapsedWeeks += weeks[idx][0] / fractionOfTeamCapacity;
    }
    elapsedWeeks = roundToTenths(elapsedWeeks);
    sims.push(elapsedWeeks);
    simsCount++;
  }

  //switching to Simulations sheet to build the chart
  ss.setActiveSheet(ss.getSheetByName("Simulations"));

  simsSheet.deleteColumns(1, 9999);
  simsSheet.appendRow(sims);

  //build a new chart representation of the Simulation
  function newChart() {
    var range = simsSheet.getRange(1, 1, 1, 10000).getValues();

    var scratchData = ArrayLib.transpose(range);
    var numRows = scratchData.length;
    var numCols = scratchData[0].length;
    var scratchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scratch");

    scratchSheet.getRange(1, 1, numRows, numCols).setValues(scratchData );
    SpreadsheetApp.flush();

    var sheet = SpreadsheetApp.getActiveSheet();
    var chartBuilder = sheet.newChart();
    chartBuilder.addRange(scratchSheet.getDataRange())
       .setChartType(Charts.ChartType.HISTOGRAM)
       .setOption('title', 'Simulations by week')
       .setOption('histogram.bucketSize', 0.2);
    sheet.insertChart(chartBuilder.setPosition(3, 1, 1, 1).build());
  }


//function to calculate the percentile
  function percentile(){

  //funktion to sort array of numbers in ascending order
  function sortNumber(a, b)
  {
    return a - b;
  }

  //defining variables for getting and setting numbers for calculation
   var tenPercent = dataSheet.getRange("J16").getValue()/ 100;
   var twentyfivePercent = dataSheet.getRange("J17").getValue()/100;
   var fiftyPercent = dataSheet.getRange("J18").getValue()/100;
   var seventyfivePercent = dataSheet.getRange("J19").getValue()/100;
   var nintyPercent = dataSheet.getRange("J20").getValue()/ 100;
   var values = simsSheet.getRange("A1:NTP1").getValues();
   var sorted = values[0].sort(sortNumber);
   var percentage = [tenPercent, twentyfivePercent, fiftyPercent, seventyfivePercent, nintyPercent];

  //calculat percentile for every percentile in percentage array
   for(i=0; i < percentage.length; i++){
     var index = sorted.length * percentage[i];

     //writting percentile in cells from datasheet
     if(percentage[i] == percentage[0]){
       dataSheet.getRange("K16").setValue(sorted[index]);
     }
     else if(percentage[i] == percentage[1]){
       dataSheet.getRange("K17").setValue(sorted[index]);
     }
     else if(percentage[i] == percentage[2]){
       dataSheet.getRange("K18").setValue(sorted[index]);
     }
     else if(percentage[i] == percentage[3]){
       dataSheet.getRange("K19").setValue(sorted[index]);
     }
     else if(percentage[i] == percentage[4]){
       dataSheet.getRange("K20").setValue(sorted[index]);
     }
    }
  }
  newChart();
  percentile();

}
