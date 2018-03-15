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
  var targetStories = dataSheet.getRange("K12").getValue();
  var fractionOfTeamCapacity = dataSheet.getRange("K13").getValue();

  var deliveredStories = 0;
  var elapsedWeeks = 0;

  var sims = [];
  var simsCount = 0;
  var simsTotal = 10000;

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

//function to calcult the percentile
function percentile(){

  //funktion to sort array of numbers in ascending order
  function sortNumber(a, b)
  {
    return a - b;
  }
  //defining variables for getting and setting numbers for calculation
   var getPercentile = simsSheet.getRange("I4").setValue("Percentile:");
   var tenPercent = dataSheet.getRange("J16").getValue()/ 100;
   var twentyfivePercent = dataSheet.getRange("J17").getValue()/100;
   var fiftyPercent = dataSheet.getRange("J18").getValue()/100;
   var seventyfivePercent = dataSheet.getRange("J19").getValue()/100;
   var ninteyPercent = dataSheet.getRange("J20").getValue()/ 100;
   var values = simsSheet.getRange("A1:NTP1").getValues();
   var sorted = values[0].sort(sortNumber);
   var percentage = [tenPercent, twentyfivePercent, fiftyPercent, seventyfivePercent, ninteyPercent];

  //calculat percentile for every percentile in percentage array
   for(i=0; i < percentage.length; i++){

     var index = sorted.length * percentage[i];

     //writting percentile in cells from datasheet
     if(percentage[i] == percentage[0]){
       var ten = dataSheet.getRange("K16").setValue(sorted[index]);
     }
     else if(percentage[i] == percentage[1]){
       var twentyfive = dataSheet.getRange("K17").setValue(sorted[index]);
     }
     else if(percentage[i] == percentage[2]){
       var fifty = dataSheet.getRange("K18").setValue(sorted[index]);
     }
     else if(percentage[i] == percentage[3]){
       var seventyfive = dataSheet.getRange("K19").setValue(sorted[index]);
     }
     else if(percentage[i] == percentage[4]){
       var nintey = dataSheet.getRange("K20").setValue(sorted[index]);
     }
   }
 }

  //calculating estimated date when x stories are finished
  function calculateDate(){

  //formating date
  function formattedDate(date){
     var year = date.getFullYear();
     var month = date.getMonth();
     var day = date.getDate();
     if(day < 10){
       day = '0'+ day;
     }
     if(month < 10){
       month = '0'+ month;
     }
     date = month+'/'+day+'/'+year;
     return date;
   }

  var weeks = dataSheet.getRange("K16:K20").getValues();

  //writting dates into datasheet
  for(i=0; i < weeks.length; i++){

    var today = new Date();

    if(weeks[i] == weeks[0]){
      var days = dataSheet.getRange("K16").getValue()*7;
      today.setDate(today.getDate()+days);
      var newDate = dataSheet.getRange("L16").setValue(formattedDate(today));
    }
    else if(weeks[i] == weeks[1]){
      var days = dataSheet.getRange("K17").getValue()*7;
      today.setDate(today.getDate()+days);
      var newDate = dataSheet.getRange("L17").setValue(formattedDate(today));
    }
    else if(weeks[i] == weeks[2]){
      var days = dataSheet.getRange("K18").getValue()*7;
      today.setDate(today.getDate()+days);
      var newDate = dataSheet.getRange("L18").setValue(formattedDate(today));
    }
    else if(weeks[i] == weeks[3]){
      var days = dataSheet.getRange("K19").getValue()*7;
      today.setDate(today.getDate()+days);
      var newDate = dataSheet.getRange("L19").setValue(formattedDate(today));
    }
    else if(weeks[i] == weeks[4]){
      var days = dataSheet.getRange("K20").getValue()*7;
      today.setDate(today.getDate()+days);
      var newDate = dataSheet.getRange("L20").setValue(formattedDate(today));
    }
  }
}

 newChart();
 percentile();
 calculateDate();
}
