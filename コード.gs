// @ts-nocheck
function calculateSalary() {
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
  var hourlyRate = settingsSheet.getRange("B2").getValue(); // 通常の時給
  var nightRate = settingsSheet.getRange("B3").getValue(); // 深夜時給
  var transportation = settingsSheet.getRange("B4").getValue(); // 交通費
  var eventName = settingsSheet.getRange("B5").getValue(); // アルバイトのイベント名
  var calendar = CalendarApp.getDefaultCalendar();
  startDate = new Date('2020-01-01');
  endDate = new Date();
  endDate.setMonth(endDate.getMonth() + 1)
  var allEvents = calendar.getEvents(startDate, endDate, {search: eventName});

  var eventsByYear = {};
  allEvents.forEach(function(event) {
    if (/^\d+-\d+(\.\d+)?$/.test(title)) {
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
        var start = event.getStartTime();
        var end = event.getEndTime();
        if (start.getFullYear() == year && start.getMonth() == month) {
          var date = start.toDateString();
          totalDays.add(date);

          var durationMinutes = (end - start) / (1000 * 60);
          var breakTime = parseInt(event.getDescription()) || 0;
          totalBreakHours += breakTime/60;
          var durationHours = durationMinutes / 60;

          while (durationHours > 0) {
            var hour = start.getHours();
            if (hour >= 22 || hour < 5) {
              nightHours += durationHours >= 1 ? 1 : durationHours;
            } else {
              dayHours += durationHours >= 1 ? 1 : durationHours;
            }
            durationHours--;
            start.setHours(start.getHours() + 1);
          }
        }
      });

      var totalPay = ((dayHours - totalBreakHours) * hourlyRate) + (nightHours * nightRate) + (totalDays.size * transportation);
      var totalHours = dayHours + nightHours - totalBreakHours
      salarySheet.getRange("B" + (month + 3)).setValue(Math.round(totalPay)); // 各月の給料を出力    
      salarySheet.getRange("C" + (month + 3)).setValue(totalHours.toFixed(2)); //勤務時間を出力
      salarySheet.getRange('D' + (month + 3)).setValue(nightHours.toFixed(2)); //深夜勤務時間を出力
      salarySheet.getRange('E' + (month + 3)).setValue(Math.round(totalDays.size)); //勤務日数を出力

    }
  }
}
