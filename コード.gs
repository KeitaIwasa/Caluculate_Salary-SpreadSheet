// @ts-nocheck
function calculateSalary() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
  var hourlyRate = settingsSheet.getRange("B2").getValue(); // 通常の時給
  var nightRate = settingsSheet.getRange("B3").getValue(); // 深夜時給
  var transportation = settingsSheet.getRange("B4").getValue(); // 交通費
  var calendar = CalendarApp.getCalendarById('emailadress@gmail.com');
  startDate = new Date('2020-01-01');
  endDate = new Date();
  endDate.setMonth(endDate.getMonth() + 1)
  var allEvents = calendar.getEvents(startDate, endDate, {search: '-'});

  var eventsByYear = {};
  allEvents.forEach(function(event) {
    if (/\d+(\.\d+)?-\d+(\.\d+)?/.test(event.getTitle()) || /\d+(\.\d+)?-L/.test(event.getTitle())) {
      var year = event.getStartTime().getFullYear();
      if (!eventsByYear[year]) {
        eventsByYear[year] = [];
      }
      eventsByYear[year].push(event);
    }
  });

  for (var year in eventsByYear) {
    var salarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(year);
    if (!salarySheet) {
      var templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template");
      salarySheet = templateSheet.copyTo(SpreadsheetApp.getActiveSpreadsheet());
      salarySheet.setName(year);
      salarySheet.getRange("A1").setValue(year); // A1セルに年を設定
      salarySheet.showSheet();
    }
    for (var month = 0; month < 12; month++) {
      var dayHours = 0;
      var nightHours = 0;
      var totalDays = new Set();
      var totalBreakHours = 0;

      eventsByYear[year].forEach(function(event) {
        var eventHours = event.getTitle().match(/\d+(\.\d+)?/g);
        var start = parseFloat(eventHours[0]);
        var startTimestamp = event.getStartTime();
        if (/\d+(\.\d+)?-\d+(\.\d+)?/.test(event.getTitle())) {
          var end = parseFloat(eventHours[1]);
        } else if (/\d+(\.\d+)?-L/.test(event.getTitle())) {
          var end = 23.5;
        }
        
        if (startTimestamp.getFullYear() == year && startTimestamp.getMonth() == month) {
          var date = startTimestamp.toDateString();
          totalDays.add(date);

          var breakTime = parseInt(event.getDescription()) || 0;
          if (!event.getDescription() && end-start>8){
            breakTime = 60
            event.setDescription(breakTime)
          } else if (!event.getDescription() && end-start>=7) {
            breakTime = 45
            event.setDescription(breakTime)
          } else if (!event.getDescription() && end-start>=6) {
            breakTime = 30
            event.setDescription(breakTime)
          }
          totalBreakHours += breakTime/60;

          if (end > 22) {
            nightHours = nightHours + (end-22);
            dayHours = dayHours + (22-start);
          } else {
            dayHours = dayHours + (end - start);
          }
        }
      });

      var totalPay = ((dayHours - totalBreakHours) * hourlyRate) + (nightHours * nightRate) + (totalDays.size * transportation);
      var totalHours = dayHours + nightHours - totalBreakHours
      salarySheet.getRange("B" + (month + 3)).setValue(Math.round(totalPay)); // 各月の給料を出力    
      salarySheet.getRange("C" + (month + 3)).setValue(totalHours.toFixed(2)); //勤務時間を出力
      salarySheet.getRange('D' + (month + 3)).setValue(nightHours.toFixed(2)); //深夜勤務時間を出力
      salarySheet.getRange('E' + (month + 3)).setValue(Math.round(totalDays.size)); //勤務日数を出力
      salarySheet.getRange('F' + (month + 3)).setValue(totalBreakHours.toFixed(2)); //休憩時間

    }
  }
}
