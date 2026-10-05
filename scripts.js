function formatTimes(timeString) {
    const formattedTimeArray = [];
  
    for (let i = 0; i < timeString.length; i += 5) {
      const timePair = timeString.substr(i, 5);
      const formattedTime = timePair.substr(0, 2) + timePair.substr(2);
      formattedTimeArray.push(formattedTime);
    }
    return formattedTimeArray;
  }
  
  function calculateTimeDifferenceInMinutes(time1, time2) {
    const [hours1, minutes1] = time1.split(':').map(Number);
    const [hours2, minutes2] = time2.split(':').map(Number);
  
    const totalMinutes1 = hours1 * 60 + minutes1;
    const totalMinutes2 = hours2 * 60 + minutes2;
  
    const differenceInMinutes = Math.abs(totalMinutes2 - totalMinutes1);
    return differenceInMinutes;
  }
  
  function processTimes(timeArray) {
    let prevTime = timeArray[0];
    let workTime = 0;
    let breakTime = 0;
    let timeDiff = 0;
    let timeCounter = 1;
    for (let i = 1; i < timeArray.length; i++) {
      if (prevTime === timeArray[i]) { // double scan
        continue;
      }
  
      if (timeCounter % 2 == 0) { // IN
        timeDiff = calculateTimeDifferenceInMinutes(prevTime, timeArray[i]);
        breakTime += timeDiff;
      } else { // OUT
        timeDiff = calculateTimeDifferenceInMinutes(prevTime, timeArray[i]);
        workTime += timeDiff;
      }
      timeCounter++;
      prevTime = timeArray[i];
    }
    console.log('Breaktime: ' + breakTime + ' mins.');
    console.log('Worktime: ' + workTime + ' mins.');
    return [workTime, breakTime];
  }
  
  function processDays(dates, daysArray) {
    let computedWorkTime = [];
    let computedBreakTime = [];
    if (daysArray.length) {
      let day;
      for (let i = 0; i < dates.length; i++) {
        day = daysArray[i]
        if (day) {
          console.log('Day ' + dates[i] + ' times: ' + JSON.stringify(formatTimes(daysArray[i])));
          timeLogs = formatTimes(daysArray[i]);
  
          if (timeLogs.length % 2) { // if time logs does not have both in and out
            computedWorkTime.push(null);
            computedBreakTime.push(null);
            continue;
          }
  
          let [workTime, breakTime] = processTimes(timeLogs);
          computedWorkTime.push(workTime);
          computedBreakTime.push(breakTime);
        } else {
          console.log('Day ' + dates[i] + ' absent.');
          computedWorkTime.push(0);
          computedBreakTime.push(0);
        }
      }
    } else {
      console.log('Completely No Data.');
    }
    return [computedWorkTime, computedBreakTime];
  }
  
  function createPersonSummary(name, dates, holidays, computedWorkTime, computedBreakTime) {
    let summary = '<tr><td colspan="4"></td></tr>';
    summary += '<tr class="name-row"><td><b>' + name + '</b></td><td><b>Work Time(hrs)</b></td><td><b>Overtime(hrs)</b></td><td><b>Break Time(hrs)</b></td></tr>';
    if (computedWorkTime.length == 0) {
      summary += '<tr><td colspan="4">Absent whole duration.</td></tr>'
      return summary;
    }
  
    const regularWorkMins = 480;
    let workHours = 0;
    let breakHours = 0;
    let overHours = 0;
    let totalWorkHours = 0;
    let totalOverHours = 0;
    let totalHolidayWorkHours = 0;
    let totalHolidayOverHours = 0;
    for (let i = 0; i < dates.length; i++) {
      overHours = 0;
      workHours = 0;
      breakHours = 0;
      let rowStyle = holidays.includes(dates[i]) ? 'background-color:lightyellow;' : '';
      if (computedWorkTime[i] == null && computedBreakTime[i] == null) {
        summary += '<tr style="' +rowStyle + '"><td>Day ' + dates[i] + '</td>' + '<td style="color:#FF0000" colspan="3"><b>incomplete time log.</b></td></tr>';
        continue;
      }
  
      if (computedWorkTime[i] > regularWorkMins) {
        workHours = regularWorkMins / 60;
        overHours = (computedWorkTime[i] - regularWorkMins) / 60;
      } else {
        workHours = computedWorkTime[i] / 60;
      }
      breakHours = computedBreakTime[i] / 60;
      if (workHours == 0 && overHours === 0 && breakHours === 0) {
        summary += '<tr style="' +rowStyle + '"><td>Day ' + dates[i] + '</td><td style="color:#0000ff" colspan="3"><b>Absent</b></td></tr>';
      } else {
        summary += '<tr style="' +rowStyle + '"><td>Day ' + dates[i] + '</td><td>' + workHours.toFixed(2) + '</td><td>' + overHours.toFixed(2) + '</td><td>' + breakHours.toFixed(2) + '</td></tr>';
      }

      // add to totals
      if (holidays.includes(dates[i])){
        totalHolidayWorkHours += workHours;
        totalHolidayOverHours += overHours;
      } else {
        totalWorkHours += workHours;
        totalOverHours += overHours;
      }
    }
    summary += '<tr><td><b>Normal Total:</b></td>' + '<td><b>' + totalWorkHours.toFixed(2) + '</b></td><td><b>' + totalOverHours.toFixed(2) + '</b></td></tr>';
    summary += '<tr><td><b>Holiday Total:</b></td>' + '<td><b>' + totalHolidayWorkHours.toFixed(2) + '</b></td><td><b>' + totalHolidayOverHours.toFixed(2) + '</b></td></tr>';
    return summary;
  }
  
  function readXLS(file, holidays) {
    const reader = new FileReader();
  
    reader.onload = function (event) {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = 'Att.log report';
      const sheet = workbook.Sheets[sheetName];
  
      const raw_data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
      const dates = raw_data[3];
      console.log('Dates:' + JSON.stringify(dates));
  
      let summary = '<table align="center">';
      let name = '';
      let timeLogsRow;
      for (let i = 4; i < raw_data.length; i++) {
        name = raw_data[i][10]
        console.log('Name:' + name);
        i++;
  
        timeLogsRow = raw_data[i];
  
        let [computedWorkTime, computedBreakTime] = processDays(dates, timeLogsRow);
        // console.log('computedWorkTime:' + JSON.stringify(computedWorkTime));
        // console.log('computedBreakTime:' + JSON.stringify(computedBreakTime));
        summary += createPersonSummary(name, dates, holidays, computedWorkTime, computedBreakTime);
      }
      summary += "</table>"
      document.getElementById('tableContainer').innerHTML = summary;
    };
  
    reader.readAsArrayBuffer(file);
  }

  function convertToArray(inputString) {
    if (inputString.trim() === "") {
        return []; // Return an empty array if the string is empty or only contains whitespace
    }

    return inputString.split(',').map(Number); // Split the string by commas and convert each part to a number
}
  
  document.getElementById('xlsFile').addEventListener('change', function (event) {
    const file = event.target.files[0];
    // Retrieve the custom text
    var textHolidays = '';
    var textHolidaysElement = document.getElementById('holidays')
    if (textHolidaysElement) {
        textHolidays = textHolidaysElement.value;
        // Proceed with processing textHolidays
        console.log(textHolidays);
    } else {
        console.error('Element with ID "holidays" not found.');
    }
    var holidays = convertToArray(textHolidays)
    if (file) {
      readXLS(file, holidays);
    }
  });
  
