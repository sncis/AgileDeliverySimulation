function monteCarlo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var simsSheet = ss.getSheets()[2];
  
  var allCharts = simsSheet.getCharts();
  for (var i in allCharts) {
    simsSheet.removeChart(allCharts[i]);
  }
  simsSheet.clear();
  
  // Todo: parameterize range inputs
  var storiesRange = dataSheet.getRange("F3:F32"); 
  var weeksRange = dataSheet.getRange("G3:G32"); 
  
  var stories = storiesRange.getValues();
  var weeks = weeksRange.getValues();
  
  // Todo: parameterize limit input
  var targetStories = dataSheet.getRange("K11").getValue();
  var fractionOfTeamCapacity = 0.5;
  
  var deliveredStories = 0;
  var elapsedWeeks = 0;
  
  var sims = [];
  var simsCount = 0;
  var simsTotal = dataSheet.getRange("K12").getValue();
  
  function roundToTenths(n) {
    n = n * 10;
    n = Math.floor(n);
    n = n / 10;
    return n;
  }

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
  
  while (simsCount < simsTotal) {
    deliveredStories = 0;
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

  ss.setActiveSheet(ss.getSheetByName("Simulations"));
  simsSheet.deleteColumns(1, 9999);
  
  simsSheet.appendRow(sims);
  
  function newChart() {
   var range = simsSheet.getRange(1, 1, 1, 10000).getValues();
   
   var scratchData = ArrayLib.transpose(range);
   var numRows = scratchData.length;
   Logger.log(scratchData.length);
   var numCols = scratchData[0].length;
   var scratchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scratch");
   scratchSheet.getRange(1, 1, numRows, numCols).setValues( scratchData );
   SpreadsheetApp.flush();
    
   var sheet = SpreadsheetApp.getActiveSheet();
   var chartBuilder = sheet.newChart();
   chartBuilder.addRange(scratchSheet.getDataRange())
       .setChartType(Charts.ChartType.HISTOGRAM)
       .setOption('title', 'Simulations by week')
       .setOption('histogram.bucketSize', 0.2);
   sheet.insertChart(chartBuilder.setPosition(3, 1, 1, 1).build());
 }
  newChart();
}