var booking_calendar = [null, null];

var formId = '1ZEj_fJDaQ54DqoePsao2VRAabEvmMfuMOduBPRFwjKY';
var form = FormApp.openById(formId);
var ssId=form.getDestinationId();
var ss = SpreadsheetApp.openById(ssId);
var sheet = ss.getActiveSheet();

function getDate(input) {
  var date = new Date();
   /* Date comes in as : MM-DD-YYYY HH:MM
      split it first on '-', then on space, and
      lastly on ':' to get Hours and Minutes*/     
  _input = input.split('-');
  date.setFullYear(_input[0]);
  date.setMonth((+_input[1]) - 1);
  _dayAndTime = _input[2].split(' ');
  _day = _dayAndTime[0]
  date.setDate(_day);
  
  _time = _dayAndTime[1];
  _hour = _time.split(':')[0];
  _min = _time.split(':')[1];
  date.setHours(_hour);
  date.setMinutes(_min);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

//find conflicts, if any, for a one-off recurring event.
function findConflictOneShot(cal, sd, ed, startTime, endTime) {
  var _existingEvents;
  _duration = endTime - startTime;
  var _startDate = new Date();
  
  var _endDate = new Date()

  for (var _c = 0; _c < cal.length; _c++) {
    _startDate.setTime(startTime);
    _endDate.setTime(endTime);
    _existingEvents = cal[_c].getEvents(sd, ed);
    for (var _i = 0; _i < _existingEvents.length; _i++) {
      _es = _existingEvents[_i].getStartTime();
      _ee = _existingEvents[_i].getEndTime();
      Logger.log("comp " + _ee.toLocaleDateString() + " | " + _endDate.toLocaleDateString() + " |" + sd.toLocaleDateString() + " " + ed.toLocaleDateString());
      if ((_startDate.valueOf() >= _es.valueOf() &&
          _startDate.valueOf() < _ee.valueOf()) ||
          (_endDate.valueOf() > _es.valueOf() &&
          _endDate.valueOf() <= _ee.valueOf()) ||
          (_startDate.valueOf() <= _es.valueOf() &&
           _endDate.valueOf() >= _ee.valueOf())) {
        Logger.log("Cal : " + _c +  " New : " + _startDate.toLocaleString() + " | " + _endDate.toLocaleString() + " | "+ " conflicts with existing " 
          + _es.toLocaleString() + " | " + _ee.toLocaleString() + " |")
        return true;
      }
    }
  }
  return false;
}

//find conflicts, if any, for a daily recurring event.
function findConflictDaily(cal, sd, ed, startTime, endTime, freq) {
  var _existingEvents;
  var _startDate = new Date();
  var _endDate = new Date();
  var _duration = endTime - startTime;
  
  for (var _c = 0; _c < cal.length; _c++) {
    _startDate.setTime(startTime);
    _endDate.setTime(endTime);

    _existingEvents = cal[_c].getEvents(sd, ed); 
    Logger.log("eventdates for scanning : " + sd + " " + ed +" " + _existingEvents.length)

    for (var _i = 0; _i < _existingEvents.length; _i++) {
      _es = _existingEvents[_i].getStartTime();
      _ee = _existingEvents[_i].getEndTime();

      if (ed.valueOf() < _es.valueOf())
        return false;

      /* We are sure that the startDates and the events will start in either the same month or the returned events will have a month
        greater than the starting event */
      if (_es.getFullYear() == _startDate.getFullYear()) {
        while(_es.getMonth() > _startDate.getMonth()) {
          _startDate.setDate(_startDate.getDate() + freq);
          _endDate.setTime(_startDate.getTime() + _duration);
        }
        
        if (_es.getFullYear() < _startDate.getFullYear() || _es.getMonth() < _startDate.getMonth())
            continue;
      
        if (_es.getDate() < _startDate.getDate()) {
          continue;
        } else if (_es.getDate() > _startDate.getDate()) {      
          while(_es.getMonth() == _startDate.getMonth() && _startDate.getDate() != _es.getDate()) {
            _startDate.setDate(_startDate.getDate() + freq);
            _endDate.setTime(_startDate.getTime() + _duration);
          }
        }
        //Logger.log("comp " + _ee.toLocaleString() + " | " + _endDate.toLocaleString() + " |")
        if ((_startDate.valueOf() >= _es.valueOf() &&
            _startDate.valueOf() < _ee.valueOf()) ||
            (_endDate.valueOf() > _es.valueOf() &&
            _endDate.valueOf() <= _ee.valueOf()) ||
            (_startDate.valueOf() <= _es.valueOf() &&
            _endDate.valueOf() >= _ee.valueOf())) {
            Logger.log("Cal :" + _c + " New : " + _startDate.toLocaleString() + " | " + _endDate.toLocaleString() + " | "+ " conflicts with existing " + _es.toLocaleString() + " | " + _ee.toLocaleString() + " |")
          return true;
        }
      } else if (_es.getFullYear() > _startDate.getFullYear()) {
        while(_es.getFullYear() > _startDate.getFullYear()) {
          _startDate.setDate(_startDate.getDate() + freq);
        }
      } else {
        continue;
      }
    }
  }
  return false;
}

//find conflicts, if any, for a weekly recurring event.
function findConflictWeekly(cal, sd, ed, startTime, endTime, daysOfWeek, startIdx, freq) {
  var _existingEvents;
  var _startDate = new Date();
  var _endDate = new Date();
  var _duration = endTime - startTime;

  var _dayIdx = startIdx;    

  var _days = [
    "MONDAY",
    "TUESDAY",
    "WEDNESDAY",
    "THURSDAY",
    "FRIDAY",
    "SATURDAY",
    "SUNDAY"
  ];
  
  for (var _c = 0; _c < cal.length; _c++) {
    _startDate.setTime(startTime);
    _endDate.setTime(endTime);

    _existingEvents = cal[_c].getEvents(sd, ed);
    Logger.log("eventdates for scanning : " + sd + " " + ed +" " + _existingEvents.length)
    for (var _i = 0; _i < _existingEvents.length; _i++) {
      _es = _existingEvents[_i].getStartTime();
      _ee = _existingEvents[_i].getEndTime();

      if (ed.valueOf() < _es.valueOf())
        return false;

      if (_es.getFullYear() == _startDate.getFullYear()) {
        if (_startDate.getMonth() != _es.getMonth()) {
          while (_startDate.getMonth() < _es.getMonth()) {
            _dayIdx = (_dayIdx + 1)%daysOfWeek.length;
            _sday = _startDate.getDay() == 0 ? 6:_startDate.getDay() - 1;
            if ((_sday) < _days.indexOf(daysOfWeek[_dayIdx]))
              _offset = _days.indexOf(daysOfWeek[_dayIdx]) - (_sday);
            else
              _offset = (7*freq) - ((_sday - 1) - _days.indexOf(daysOfWeek[_dayIdx]));
            
            _startDate.setDate(_startDate.getDate() + _offset);
            _endDate.setTime(_startDate.getTime() + _duration);
          }

          if (_es.getFullYear() < _startDate.getFullYear() || _es.getMonth() < _startDate.getMonth())
            continue;
        }

        if (_startDate.getMonth() == _es.getMonth()) {
          if (_es.getDate() < _startDate.getDate()) {
            //Logger.log("continue to next event " );
            continue;
          } else if (_es.getDate() > _startDate.getDate()) {
            while(_es.getFullYear() == _startDate.getFullYear() && _es.getMonth() == _startDate.getMonth() && _startDate.getDate() < _es.getDate()) {            
              _dayIdx = (_dayIdx + 1)%daysOfWeek.length;
              _sday = _startDate.getDay() == 0 ? 6:_startDate.getDay() - 1;
              if ((_sday) < _days.indexOf(daysOfWeek[_dayIdx]))
                _offset = _days.indexOf(daysOfWeek[_dayIdx]) - _sday;
              else
                _offset = (7*freq) - ((_sday) - _days.indexOf(daysOfWeek[_dayIdx]));
              _startDate.setDate(_startDate.getDate() + _offset);
              _endDate.setTime(_startDate.getTime() + _duration);
            }
          }
          //If the loop above caused the year or month to change
          if (_es.getFullYear() < _startDate.getFullYear() || _es.getMonth() < _startDate.getMonth())
            continue;
          //Logger.log("comp " + _es.toLocaleString() + " | " + _ee.toLocaleString() + " |" + _startDate.toLocaleString() + " | " + _endDate.toLocaleString())
          
          if ((_startDate.valueOf() >= _es.valueOf() &&
              _startDate.valueOf() < _ee.valueOf()) ||
              (_endDate.valueOf() > _es.valueOf() &&
              _endDate.valueOf() <= _ee.valueOf()) ||
              (_startDate.valueOf() <= _es.valueOf() &&
              _endDate.valueOf() >= _ee.valueOf())) {
            Logger.log("Cal :" + _c + " New : " + _startDate.toLocaleString() + " | " + _endDate.toLocaleString() + " | "+ " conflicts with existing " + _es.toLocaleString() + " | " + _ee.toLocaleString() + " |")
            
            return true;
          }
        }
      } else if (_es.getFullYear() > _startDate.getFullYear()){
        while(_es.getFullYear() > _startDate.getFullYear()) {
          _dayIdx = (_dayIdx + 1)%daysOfWeek.length;
          _sday = _startDate.getDay() == 0 ? 6:_startDate.getDay() - 1;
          if ((_sday) < _days.indexOf(daysOfWeek[_dayIdx]))
            _offset = _days.indexOf(daysOfWeek[_dayIdx]) - (_sday);
          else
            _offset = (7*freq) - ((_sday) - _days.indexOf(daysOfWeek[_dayIdx]));

          _startDate.setDate(_startDate.getDate() + _offset);
          _endDate.setTime(_startDate.getTime() + _duration);
        }

      } else {
        continue;
      }
    }
  }
  return false;
}
function getNextDateMonthly(startFirstDate, curDate, sameDateEveryMonth, freq) {
  return getLastDateMonthly(startFirstDate, curDate, sameDateEveryMonth, freq, 2);
}

function findConflictMonthly(cal, sd, ed, startTime, endTime, sameDateEveryMonth, freq) {
  var _existingEvents;
  var _duration = endTime - startTime;
  var _startDate = new Date();
  var _endDate = new Date();
  var _tempStartDate = new Date();
  for (var _c = 0; _c < cal.length; _c++) {
    _startDate.setTime(startTime);
    _endDate.setTime(endTime);
    _tempStartDate.setTime(startTime);
    
    _existingEvents = cal[_c].getEvents(sd, ed);
    
    Logger.log("eventdates for scanning : " + sd + " " + ed +" " + _existingEvents.length)
  
    for (var _i = 0; _i < _existingEvents.length; _i++) {
      _es = _existingEvents[_i].getStartTime();
      _ee = _existingEvents[_i].getEndTime();
      if (ed.valueOf() < _es.valueOf()) {
        return false;
      }
      
      if (_es.getFullYear() == _startDate.getFullYear()) {
        if (_es.getMonth() > _startDate.getMonth()) {
          while(_es.getMonth() > _startDate.getMonth()) {
            _startDate = getNextDateMonthly(sd, _startDate, sameDateEveryMonth, freq);
          }
          _endDate.setTime(_startDate.getTime() + _duration);
        }
        if (_es.getFullYear() < _startDate.getFullYear() || _es.getMonth() < _startDate.getMonth())
          continue;
        
        if (_ee.getDate() < _startDate.getDate()) {
          continue;
        } else if (_es.getDate() > _startDate.getDate()) {
          _startDate = getNextDateMonthly(sd, _startDate, sameDateEveryMonth, freq);
          _endDate.setTime(_startDate.getTime() + _duration);
        }
    
        Logger.log("comp " + _ee.toLocaleString() + " | " + _startDate.toLocaleString() + " | ED " +  _endDate.toLocaleString())
        if ((_startDate.valueOf() >= _es.valueOf() &&
            _startDate.valueOf() < _ee.valueOf()) ||
            (_endDate.valueOf() > _es.valueOf() &&
            _endDate.valueOf() <= _ee.valueOf()) ||
            (_startDate.valueOf() <= _es.valueOf() &&
            _endDate.valueOf() >= _ee.valueOf())) {
          Logger.log("Cal :" + _c + " New : " + _startDate.toLocaleString() + " | " + _endDate.toLocaleString() + " | "+ " conflicts with existing " + _es.toLocaleString() + " | " + _ee.toLocaleString() + " |")
          return true;
        }  
      } else if (_es.getFullYear() > _startDate.getFullYear()) {
        while (_startDate.getFullYear() < _es.getFullYear()) {
          _startDate.setMonth(_startDate.getMonth() + freq);
        }

        if (!sameDateEveryMonth) {
          _tempStartDate = getNextDateMonthly(sd, _startDate, sameDateEveryMonth, freq);
        }
        _endDate.setTime(_startDate.getTime() + _duration);
      } else {
        continue;
      }  
    }
  }
  return false;
}

function getLastDateMonthly(startFirstDate, curDate, sameDateEveryMonth, freq, occurencesDays) { 
  var le = new Date();
  if (curDate == null)
    _month = startFirstDate.getMonth();
  else
    _month = curDate.getMonth();

  _date = startFirstDate.getDate();
  _numMonths = ((occurencesDays -1) * freq);
  _lastMonth = ((_month + _numMonths)%12);
   
  if (curDate == null) {
    _yr = startFirstDate.getFullYear();
    
    _numYears = Math.floor(_numMonths/12);
    
    le.setFullYear(_numYears + startFirstDate.getFullYear());
  } else {
    _yr = curDate.getFullYear();
    le.setFullYear(_yr);
  }
  
  if (_lastMonth <= _month)
    le.setFullYear(1 + le.getFullYear());
      
  le.setMonth(_lastMonth);
  
  if (sameDateEveryMonth) {
    if (_date < 29) {
      
      le.setDate(_date);
    } else {
      months30 = [
        0, //CalendarApp.Month.JANUARY,
        2, //CalendarApp.Month.MARCH,
        3, //CalendarApp.Month.APRIL,
        4, //CalendarApp.Month.MAY,
        5, //CalendarApp.Month.JUNE,
        6, //CalendarApp.Month.JULY,
        7, //CalendarApp.Month.AUGUST,
        8, //CalendarApp.Month.SEPTEMBER,
        9, //CalendarApp.Month.OCTOBER,
        10,//CalendarApp.Month.NOVEMBER,
        11,//CalendarApp.Month.DECEMBER
      ];
      months31 = [
        0, //CalendarApp.Month.JANUARY,
        2, //CalendarApp.Month.MARCH,
        4, //CalendarApp.Month.MAY,
        6, //CalendarApp.Month.JULY,
        7, //CalendarApp.Month.AUGUST,
        9, //CalendarApp.Month.OCTOBER,
        11,//CalendarApp.Month.DECEMBER
      ];
      _nextMonth = _month;
      _ml = [months30, months31];
      _mlIdx = _date%10;
      
      //Skip over months that don't have this date
      for (var _i = 0; _i < occurencesDays - 1; _i++) {
        var _prevMonth = _nextMonth;
        _nextMonth = (_prevMonth + freq) % 12;
        
        if (_date == 29 && _nextMonth == 1 && _yr/4 != 0) {
          _nextMonth+=freq;
        } else if (_ml[_mlIdx].includes(_nextMonth) == false) {
          while(_ml[_mlIdx].includes(_nextMonth) == false)
            _nextMonth+=freq;
        }
        if (_nextMonth < _prevMonth)
          _yr+=1;
        Logger.log(_nextMonth)
      }
      le.setFullYear(_yr);
      le.setMonth(_nextMonth);
      le.setDate(_date);
    }
  } else {
    _sDay = startFirstDate.getDay();
    //Logger.log(le.toString() + " " + _sDay);
    if (_date < 29) {
      le.setDate(Math.floor((_date-1)/7)*7 + 1)
      _lastDay = le.getDay();
      if (_lastDay > _sDay)
        le.setDate (le.getDate() + 7 - (_lastDay - _sDay));
      else
        le.setDate (le.getDate() + _sDay - _lastDay);
    } else {
      _lastDate = [31,29,31,30,31,30,31,31,30,31,30,31];
      le.setDate(_lastDate[_lastMonth]);
      _lastDay = le.getDay();
      if (_lastDay >= _sDay)
        le.setDate (le.getDate() - (_lastDay - _sDay));
      else {
      //  Logger.log(le.getDate() + " " + _sDay + " " + _lastDay)
        le.setDate (le.getDate() - (7 - (_sDay - _lastDay)));
      }
    }
  }
  
  if (curDate != null) {
    le.setMinutes(curDate.getMinutes());
    le.setHours(curDate.getHours());
    le.setSeconds(0);
    le.setMilliseconds(0);
    Logger.log("LE : " + le.toLocaleString() + " CD " + curDate.toLocaleString());
  }
  return le;
}

function doCreateEvent(description, eMail, contactNum, title, startFirstDate, duration, room, recurring, freq, unit, occurencesDays, daysOfWeek, sameDateEveryMonth, lastRecurrenceDate) {
  var daysList = ["MONDAY", "TUESDAY", "WEDNESDAY","THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];
  var endFirstDate = new Date();
  _offSet = 0;
  
  var bigRoom = CalendarApp.getOwnedCalendarsByName('Big Room');
  var smallRoom = CalendarApp.getOwnedCalendarsByName('Small Room');

  Logger.log("Big Room Cal : " + bigRoom[0].getId())
  Logger.log("Small Room Cal : " + smallRoom[0].getId());
  _sub = 0;
  /* Date comes in as : MM-DD-YYYY HH:MM
      split it first on '-', then on space, and
      lastly on ':' to get Hours and Minutes*/
  /* Duration comes in as 'HH Hours MM minutes' */
  _s = duration.split(" ");
  
  if (_s.length == 6) {
      _minutes = +_s[0] * 60;
      _minutes += +_s[3];
  } else {
    if (_s[_s.length - 1] == "minutes")
      _minutes = +_s[0];
    else
      _minutes = +_s[0] * 60;
  }
  if (room == 1)
    booking_calendar = [bigRoom[0]];
  else if (room == 2)
    booking_calendar = [smallRoom[0]]
  else
    booking_calendar = [bigRoom[0], smallRoom[0]]
  
  /* This is the date used to create events, both recurring and one-shot */
  endFirstDate.setSeconds(0);
  endFirstDate.setTime(startFirstDate.getTime() + _minutes * 60000); //takes milliseconds as input

  Logger.log("get startFirstDate %s %s %s %s", duration, startFirstDate.getMonth() ,startFirstDate.toLocaleString(), endFirstDate.toLocaleString());
  
  if (recurring == "No") {
    var _baseDateOneOff = new Date();
    _baseDateOneOff.setTime(startFirstDate.getTime())
    _baseDateOneOff.setHours(0);
    _baseDateOneOff.setMinutes(0);
    _baseDateOneOff.setSeconds(0)
    
    var _lastDateOneOff=new Date();
    _lastDateOneOff.setTime(endFirstDate.getTime());
    _lastDateOneOff.setHours(23);
    _lastDateOneOff.setMinutes(59);
    _lastDateOneOff.setSeconds(59);
    //Logger.log("one shot get startFirstDate %s %s %s", duration, startFirstDate.getMonth() ,startFirstDate.toLocaleString());
    //Logger.log("base  %s %s %s", duration, _baseDateOneOff.toLocaleDateString() ,_lastDateOneOff.toLocaleTimeString());
    //Check if there is a conflict. If both rooms have been requested for then we'll search for a conflict in both calendars
    if (findConflictOneShot(booking_calendar, _baseDateOneOff, _lastDateOneOff, startFirstDate.getTime(), endFirstDate.getTime()) == true)
		  return null;
    var _descr = description?description + " " : title + " ";
    if (contactNum)
      _descr+= "\nContact No. : " + contactNum;
	  var event1 = booking_calendar[0].createEvent(title, startFirstDate, endFirstDate, {description: _descr, guests:eMail});
    if (event1 == null)
      return null;
  
    if (room == 1) {
      event1.setLocation("Big Room")
      Logger.log("Location : " + event1.getLocation())
      return [event1.getId(), -1];
    } else if (room == 2) {
      event1.setLocation("Small Room");
      Logger.log("Location : " + event1.getLocation())
      return [-1, event1.getId()];
    } else {
      var event2 = booking_calendar[1].createEvent(title, startFirstDate, endFirstDate, {description: _descr, guests:eMail});
      if (event2 == null) {
        event1.deleteEventSeries();
        return null;
      }
    }
    event1.setLocation("Both Rooms");
    event2.setLocation("Both Rooms");
    event1.setColor(CalendarApp.EventColor.MAUVE);
    event2.setColor(CalendarApp.EventColor.MAUVE);
    Logger.log("Location : " + event1.getLocation() + " " + event2.getLocation())

	  //Logger.log("endFirstDate %s %s %s", duration, endFirstDate.toLocaleDateString() ,endFirstDate.toLocaleTimeString());
	  return [event1.getId(), event2.getId()];
  }

  var evRecurrence = CalendarApp.newRecurrence();
  var eventSeriesIds = [-1, -1];
  switch(unit) {
    case "Day(s)":
      _daysOffset = freq * (occurencesDays - 1); 
      var _baseDateD = new Date();
      _baseDateD.setTime(startFirstDate.getTime())
      _baseDateD.setHours(0);
      _baseDateD.setMinutes(0);
      _baseDateD.setSeconds(0)
      
      var _lastDateD=new Date();
      _lastDateD.setTime(endFirstDate.getTime() + (_daysOffset * 24 * 3600 * 1000));
      _lastDateD.setHours(23);
      _lastDateD.setMinutes(59);
      _lastDateD.setSeconds(59);

      //Check if there is a conflict
      if (findConflictDaily(booking_calendar, _baseDateD, _lastDateD,
				  startFirstDate.getTime(), endFirstDate.getTime(), freq) == true) 
		    return null;

      var _descr = description?description + " " : title + " ";
      if (contactNum)
        _descr+= "\nContact No. : " + contactNum;
	    var dailyRecurrence = evRecurrence.addDailyRule().times(occurencesDays).interval(freq);
      var eventSeries1 = booking_calendar[0].createEventSeries(title,  startFirstDate, endFirstDate, dailyRecurrence, {description:_descr, guests:eMail});
      if (eventSeries1 == null)
        return null;

      if (room == 1)
        eventSeriesIds = [eventSeries1.getId(), -1];
      else if (room == 2)
        eventSeriesIds = [-1, eventSeries1.getId()];
      else {
        var eventSeries2 = booking_calendar[1].createEventSeries(title,  startFirstDate, endFirstDate, dailyRecurrence, {description:_descr, guests:eMail});
        if (eventSeries2 == null) {
          eventSeries1.deleteEventSeries();
          return null;
        }
        eventSeriesIds = [eventSeries1.getId(), eventSeries2.getId()];
      }

      lastRecurrenceDate.setTime(_lastDateD.getTime());
      break;
  
    case "Week(s)":
      var _startIdx = 0;
      var _numOfWeeks = 0;
      var _lastDay = 0;
      var _firstDay = 0;
      var _firstDayGap = 0;
      var _found = 0;
      
      /*
        for (i = 0; i < daysOfWeek.length; i++)
          _days[i] = CalendarApp.Weekday[daysOfWeek[i]]
        
        The following statements achieves the same effect.
      */
      _days = daysOfWeek.map(function(day) { Logger.log(day); return CalendarApp.Weekday[day] });
      
      _firstDay = startFirstDate.getDay() == 0? 6 : startFirstDate.getDay() - 1; // Week begins on Monday
	     /*_baseDateW is the date on which the first event falls in case the
		    startFirstDate is before the first day in the list of days the event is to
		    be repeated. */
      var _baseDateW = new Date();
      _baseDateW.setTime(startFirstDate.getTime())
      _baseDateW.setHours(0);
      _baseDateW.setMinutes(0);
      _baseDateW.setSeconds(0)
      /* if the day the startDate specified falls on is not in the desired days, go to the next day available.
        e.g., if the day is Sat but the day desired are Mon, Wed, Fri, Sun, set the _firstDay as Sun*/
      Logger.log("FD : " + _firstDay)
      for (i = 0; i < daysOfWeek.length; i++){
        if (_firstDay <=  daysList.indexOf(daysOfWeek[i])) {
          _firstDayGap = daysList.indexOf(daysOfWeek[i]) - _firstDay;
          _firstDay += _firstDayGap;
          _baseDateW.setDate(startFirstDate.getDate() + _firstDayGap);
          startFirstDate.setDate(_baseDateW.getDate());
          endFirstDate.setDate(endFirstDate.getDate() + _firstDayGap);
          break;
        }
      }
      Logger.log("FD : " + _firstDay)
      if (_firstDay > daysList.indexOf(daysOfWeek[0])) {
        _found = 0;
        for (i = 1; i < daysOfWeek.length; i++) {
          if (_firstDay <= daysList.indexOf(daysOfWeek[i])) {
            _startIdx = i;
            _found = 1;
            Logger.log("start idx " + _startIdx)
            break;
          }
        }

        if (!_found) {
          _startIdx = 0;
          _baseDateW.setDate(startFirstDate.getDate() + 7 - daysList.indexOf(daysOfWeek[0]) + 1);
          startFirstDate.setDate(_baseDateW.getDate());
          endFirstDate.setDate(endFirstDate.getDate() + 7 - daysList.indexOf(daysOfWeek[0]) + 1);
        } else {
          Logger.log("bef base date : " + _baseDateW.getDate());
          _baseDateW.setDate(startFirstDate.getDate() + _firstDayGap);
          startFirstDate.setDate(_baseDateW.getDate());
          endFirstDate.setDate(endFirstDate.getDate() + _firstDayGap);
          Logger.log("base date : " + _baseDateW.getDate());
        }
      } else if (_firstDay < daysList.indexOf(daysOfWeek[0])) {
        _baseDateW.setDate(startFirstDate.getDate() + daysList.indexOf(daysOfWeek[0]) - _firstDay);
        startFirstDate.setDate(_baseDateW.getDate());
        endFirstDate.setDate(endFirstDate.getDate() + daysList.indexOf(daysOfWeek[0]) - _firstDay);
        _firstDay = daysList.indexOf(daysOfWeek[0]);
      }

      _numEventsPerWeek = daysOfWeek.length;
      _numEventsFirstWeek = _numEventsPerWeek - _startIdx;
      
      _remainingEvents = occurencesDays - _numEventsFirstWeek;
      _numOfWeeks = 1;
      
      _spillOver = (_remainingEvents)%_numEventsPerWeek;
      
      if (_spillOver > 0) {
        _lastDayIdx = _spillOver - 1;
        _numOfWeeks = _numOfWeeks + 1;
      } else 
        _lastDayIdx = _numEventsPerWeek - 1;
      

      _numOfWeeks = _numOfWeeks + Math.floor((_remainingEvents)/_numEventsPerWeek);
      //Logger.log("nw1 : " + _numOfWeeks + " " + " " + _remainingEvents + " " + _lastDayIdx + " " + Math.floor((_remainingEvents)/_numEventsPerWeek));   
      _numOfWeeks = _numOfWeeks + ((_numOfWeeks - 1) * (freq - 1)); 
      //Logger.log("nw2: " + _numOfWeeks + " " + _baseDateW.toLocaleDateString());   

      //Logger.log(_lastDay + " " + freq + " " + _numOfWeeks + " " + _numEventsPerWeek + " " + _spillOver + " " + occurencesDays)
      /* _lastDateW is the date on which the last event will fall, i.e.,
      * occurencesDays'th instance*/
      /* _lastDateW and _baseDateW help in running through the events in the calendar
      * and ascertaining if there are any conflicts. */
      var _lastDateW=new Date();
      
      _lastDateW.setTime(_baseDateW.getTime() + ((_numOfWeeks - 1) * 7 * 24 * 3600 * 1000));
      Logger.log("base date : " + _baseDateW.toString() + " " + daysList.indexOf(daysOfWeek[_lastDayIdx]));
      {

        _finalDay = (_lastDateW.getDay() == 0)? 6 : _lastDateW.getDay() - 1; //Week begins on Monday
        

        _finalDay = _finalDay - daysList.indexOf(daysOfWeek[_lastDayIdx])
        _lastDateW.setDate(_lastDateW.getDate() - _finalDay)
        Logger.log("last date : " + _lastDateW.toString() + " " + _finalDay + " " + daysList.indexOf(daysOfWeek[_lastDayIdx]));
      }

      _lastDateW.setHours(23);
      _lastDateW.setMinutes(59);
      _lastDateW.setSeconds(59);

      //Check if there is a conflict : cal, sd, ed, startTime, endTime, daysOfWeek, startIdx, freq, occurencesDays
      if (findConflictWeekly(booking_calendar, _baseDateW, _lastDateW,
				  startFirstDate.getTime(), endFirstDate.getTime(), daysOfWeek, _startIdx, freq) == true) {
		    return null;
	    }
      var _descr = description?description + " " : title + " ";
      if (contactNum)
        _descr+= "\nContact No. : " + contactNum;
	  
      var weeklyRecurrence = evRecurrence.addWeeklyRule().onlyOnWeekdays(_days).interval(freq).times(occurencesDays);
      var eventSeries1 = booking_calendar[0].createEventSeries(title,  startFirstDate, endFirstDate, weeklyRecurrence, {description:_descr, guests:eMail});
      if (eventSeries1 == null)
        return null;

      if (room == 1)
        eventSeriesIds = [eventSeries1.getId(), -1];
      else if (room == 2)
        eventSeriesIds = [-1, eventSeries1.getId()];
      else {
        var eventSeries2 = booking_calendar[1].createEventSeries(title, startFirstDate, endFirstDate, weeklyRecurrence, {description:_descr, guests:eMail});
        if (eventSeries2 == null) {
          eventSeries1.deleteEventSeries();
          return null;
        }
        eventSeriesIds = [eventSeries1.getId(), eventSeries2.getId()];
      }

      lastRecurrenceDate.setTime(_lastDateW.getTime());
      //Logger.log("index: |" + daysList.indexOf(daysOfWeek[0]) + " " + _firstDay + " " +  endFirstDate.toLocaleTimeString() + " " + _lastDateW.toLocaleDateString());
      break;

    case "Month(s)":
      var monthlyRecurrence;
      var _calDays = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31];
      var _days = [
        CalendarApp.Weekday.MONDAY,
        CalendarApp.Weekday.TUESDAY,
        CalendarApp.Weekday.WEDNESDAY,
        CalendarApp.Weekday.THURSDAY,
        CalendarApp.Weekday.FRIDAY,
        CalendarApp.Weekday.SATURDAY,
        CalendarApp.Weekday.SUNDAY,
      ];

      var _weekDay = _days[(startFirstDate.getDay() == 0)?6:startFirstDate.getDay() - 1];
      var _monthDay = startFirstDate.getDate();
      
      var _baseDateM = new Date();
      _baseDateM.setTime(startFirstDate.getTime())
      _baseDateM.setHours(0);
      _baseDateM.setMinutes(0);
      _baseDateM.setSeconds(0)
      

      var _lastDateM = getLastDateMonthly(startFirstDate, null, sameDateEveryMonth, freq, occurencesDays);
      _lastDateM.setHours(23);
      _lastDateM.setMinutes(59);
      _lastDateM.setSeconds(59)
      Logger.log("LD :" + _lastDateM.toLocaleString());
      //Logger.log(_shishir)
      //Check if there is a conflict
      if (findConflictMonthly(booking_calendar, _baseDateM, _lastDateM,
				  startFirstDate.getTime(), endFirstDate.getTime(), sameDateEveryMonth, freq) == true)
		    return null;

      //Logger.log(_shishir)

      if (sameDateEveryMonth)
        monthlyRecurrence = evRecurrence.addMonthlyRule().interval(freq).times(occurencesDays);
      else {
        var _weekStart;
        var _validDays;

        _weekStart = Math.floor((_monthDay-1)/7) * 7; // Round Down
          //Logger.log(_validDays + " " + _weekStart + " " + _monthDay)
        _validDays = _calDays.slice(_weekStart + 1, _weekStart + 1 + 7)
        
        Logger.log(_validDays + " W " + _weekDay);
        //Logger.log(_shishir)
        
        if (_monthDay < 29)
          monthlyRecurrence = evRecurrence.addMonthlyRule().onlyOnWeekday(_weekDay).onlyOnMonthDays(_validDays).interval(freq).times(occurencesDays);
        else {
          months = [
            CalendarApp.Month.JANUARY,
            CalendarApp.Month.FEBRUARY,
            CalendarApp.Month.MARCH,
            CalendarApp.Month.APRIL,
            CalendarApp.Month.MAY,
            CalendarApp.Month.JUNE,
            CalendarApp.Month.JULY,
            CalendarApp.Month.AUGUST,
            CalendarApp.Month.SEPTEMBER,
            CalendarApp.Month.OCTOBER,
            CalendarApp.Month.NOVEMBER,
            CalendarApp.Month.DECEMBER
          ];
          months30 = [
           CalendarApp.Month.APRIL,
           CalendarApp.Month.JUNE,
           CalendarApp.Month.SEPTEMBER,
           CalendarApp.Month.NOVEMBER,
          ],
          months31 = [
           CalendarApp.Month.JANUARY,
           CalendarApp.Month.MARCH,
           CalendarApp.Month.MAY,
           CalendarApp.Month.JULY,
           CalendarApp.Month.AUGUST,
           CalendarApp.Month.OCTOBER,
           CalendarApp.Month.DECEMBER
          ]
          //months of interest
          if (freq > 1) {
            _feb = false;
            _m30 = [0];
            _idx30 = 0;
            _m31 = [0];
            _idx31 = 0;
            _prevMonth = -1;
            _firstMonth = startFirstDate.getMonth();
            _nextMonth = _firstMonth;
            while(1) {
              _prevMonth = _nextMonth;
              if (months30.includes(months[_prevMonth]) == true) {
                _m30[_idx30] = months[_prevMonth];
                _idx30++; 
              } else if (months31.includes(months[_prevMonth]) == true) {
                _m31[_idx31] = months[_prevMonth];
                _idx31++;
              } else if (_prevMonth == 1)
                _feb = true;

              _nextMonth = (_prevMonth + freq) % 12;
              if (_nextMonth == _firstMonth)
                break;
            }
          }
          Logger.log(_m30 + " " + _m31)
          if (_feb == true)
            Logger.log("FEB")

          if (_feb == true)
            monthlyRecurrence = evRecurrence.addMonthlyRule().interval(freq).onlyInMonths(_m31).onlyOnWeekday(_weekDay).onlyOnMonthDays([25,26,27,28,29,30,31]).until(_lastDateM).
              addMonthlyRule().interval(freq).onlyInMonth(CalendarApp.Month.FEBRUARY).onlyOnWeekday(_weekDay).onlyOnMonthDays([22,23,24,25,26,27,28, 29]).until(_lastDateM).
              addMonthlyRule().interval(freq).onlyInMonths(_m30).onlyOnWeekday(_weekDay).onlyOnMonthDays([24, 25,26,27,28,29,30]).until(_lastDateM);
          else
            monthlyRecurrence = evRecurrence.addMonthlyRule().interval(freq).onlyInMonths(_m31).onlyOnWeekday(_weekDay).onlyOnMonthDays([25,26,27,28,29,30,31]).until(_lastDateM).
              addMonthlyRule().interval(freq).onlyInMonths(_m30).onlyOnWeekday(_weekDay).onlyOnMonthDays([24, 25,26,27,28,29,30]).until(_lastDateM);
        }
      }
      var _descr = description?description + " " : title + " ";
      if (contactNum)
        _descr+= "\nContact No. : " + contactNum;
	  
      var eventSeries1 = booking_calendar[0].createEventSeries(title,  startFirstDate, endFirstDate, monthlyRecurrence, {description:_descr, guests:eMail});
;
      if (eventSeries1 == null)
        return null;

      if (room == 1)
        eventSeriesIds = [eventSeries1.getId(), -1];
      else if (room == 2)
        eventSeriesIds = [-1, eventSeries1.getId()];
      else {
        var eventSeries2 = booking_calendar[1].createEventSeries(title, startFirstDate, endFirstDate, monthlyRecurrence, {description:_descr, guests:eMail});
        if (eventSeries2 == null) {
          eventSeries1.deleteEventSeries();
          return null;
        }
        eventSeriesIds = [eventSeries1.getId(), eventSeries2.getId()];
      }

      lastRecurrenceDate.setTime(_lastDateM.getTime());
      break;
  }
  lastRecurrenceDate.setHours(startFirstDate.getHours());
  lastRecurrenceDate.setMinutes(startFirstDate.getMinutes());
  lastRecurrenceDate.setSeconds(0);
      
  //Logger.log("recuring endFirstDate %s %s %s", duration, lastRecurrenceDate.toLocaleDateString() ,lastRecurrenceDate.toLocaleTimeString());
  if (room == 1) {
    eventSeries1.setLocation("Big Room");
  } else if (room == 2) {
    eventSeries1.setLocation("Small Room");
  } else {
    eventSeries1.setLocation("Both Rooms");
    eventSeries2.setLocation("Both Rooms");
    eventSeries1.setColor(CalendarApp.EventColor.YELLOW);
    eventSeries2.setColor(CalendarApp.EventColor.YELLOW);
  }
  if (room == 3) {
    Logger.log("Location : " + eventSeries1.getLocation() + " " + eventSeries2.getLocation())
  } else if (room == 2) {
    Logger.log("Location : " + eventSeries1.getLocation())
  } else
    Logger.log("Location : " + eventSeries1.getLocation())
  return eventSeriesIds;
}

function parseFormResponse(lastRow) {
  Logger.log('Calling the Forms API!');
  var unit = 0;
  var formId = '1ZEj_fJDaQ54DqoePsao2VRAabEvmMfuMOduBPRFwjKY';
  var form = FormApp.openById(formId);
  
  //sheet info
  var rangeData = sheet.getDataRange();

  var lastColumn = rangeData.getLastColumn();
  
  //Form Response 
  var startFirstDate;
  var evId;
  var delEventId;
  var description;
  var eMail;
  var contactNum;
  var title;
  var duration;
  var recurring;
  var freq;
  var unit;
  var occurencesDays;
  var daysOfWeek;
  var sameDateEveryMonth = 0;
  var fromGE_ = null;

  var lastRecurrenceDate = new Date();

  eMail = sheet.getRange(lastRow,2).getValue()
  Logger.log(lastColumn + " " + lastRow)
  
  
  var formResponses = form.getResponses();
  
  duration = 0, recurring = 0, daysOfWeek = 0, freq = 0;
  
  var formResponse = formResponses[formResponses.length - 1];
  var itemResponses = formResponse.getItemResponses();

  for (var j = 0; j < itemResponses.length; j++) {
	  var itemResponse = itemResponses[j];

	  switch(itemResponse.getItem().getTitle()) {
		  case "Contact Number":
			  /*
          Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
			  */
        contactNum = itemResponse.getResponse();
			  break;

		  case "Title":
			  /* 
          Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
			  */
        title = itemResponse.getResponse();
			  break;

		  case "Summary":
			  description = itemResponse.getResponse();
			  break;

		  case "From":
			  fromGE_ = itemResponse.getResponse();
			  break;

		  case "Resident's Name":
			  requester = itemResponse.getResponse();
			  break;

		  case "Action":
			  resp = itemResponse.getResponse();
			  if (resp == "Add")
				  action = 1;
			  else if (resp == "Delete")
				  action = 2;
			  break;

      case "Room":
        resp = itemResponse.getResponse();
        if (resp == "Big")
          room = 1;
        else if (resp == "Small")
          room = 2;
        else 
          room = 3;
        break;
       
		  case "Recurring":
			  recurring = itemResponse.getResponse();
			  /*
          Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
        */
			  break;

		  case "Start Date and Time":
      case "Start Date and Time (Non-recurring)":
			  resp = itemResponse.getResponse();
			  startFirstDate = getDate(resp);
			  //Logger.log("#%s startFirstDate %s %s", (1).toString(), startFirstDate.toLocaleDateString(), resp);
			  break;

		  case "Event Id":
			  delEventId = itemResponse.getResponse();
			  //Logger.log("#%s Event Id for deletion %s ", (1).toString(), delEventId);
			  break;

		  case "Duration":
      case "Duration (Non-recurring)":
			  duration = itemResponse.getResponse();

			  /* 
          Logger.log('Response to %s the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
			  */
        break;

		  case "Repeats every":
			  freq = +itemResponse.getResponse();

			  /*
          Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
        */
			  j++;
			  /* case (... repetition unit): (Day/Week/Month/Year)*/
			  itemResponse = itemResponses[j];
			  unit = itemResponse.getResponse();
			  /* 
          Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
			  */
        break;

		  case "Repeat on":
			  daysOfWeek = itemResponse.getResponse();
			  /*
          Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
        */
			  break;

		  case "Ends after (number of days)":
			  /* repetition days means number of days in total. If an event starts on 1st of a month
				 and is repeated every other day for 12 repetitions, that means there will be 12 days
				 that the event will occur on and the resource booked for. */
			  occurencesDays = +itemResponse.getResponse();
			  /*
          Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
			  */
        break;

		  case "Ends after (number of weeks)":
			  occurencesWeeks = +itemResponse.getResponse();
			  /*
          Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
			  */
        break;

		  case "On":
			  if (itemResponse.getResponse() == "Same day as start date")
				  sameDateEveryMonth = 0; 
			  else if (itemResponse.getResponse() =="Same date as start date")
				  sameDateEveryMonth = 1;
        /*
			  Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
			  */
        break;

		  case "Ends after (number of months)":
			  occurencesMonths = +itemResponse.getResponse();
			  /*
          Logger.log('Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
			  */
        break;

		  default:
			  Logger.log('Default Response #%s to the question "%s" was "%s"',
					  (1).toString(),
					  itemResponse.getItem().getTitle(),
					  itemResponse.getResponse());
			  break;
	  }
  }
  
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  
  if (!lock.hasLock()) {
    Logger.log('Could not obtain lock after 30 seconds.');
    {
      var sub = "FAILED!! Your Event Creation Request for Start Date : " + startFirstDate.toDateString() + " failed.";
      var resp = "Location : Duet :" 
      if (room == 1)
       resp += "Big Room\n"
      else if (room == 2)
        resp += "Small Room\n"
      else 
        resp+= "Both Rooms\n"
    
      resp+= "Title : " + title + "\n"
      
      if (description)
        resp +="Description : " + description + " \n"

      resp += "Start Time : " + startFirstDate.toLocaleDateString() + " " + startFirstDate.toLocaleTimeString() + "\n"
      + "Duration : " + duration + "\n"
      + "Recurring : " + recurring + "\n" 
      + "Requester's Name :" + requester + "\n"
      + "From : " + fromGE_ + "\n";
      if (contactNum)
        resp+= "Contact Number : " + contactNum + "\n";
      
      resp+="This likely happened because of a system issue. Please try again.";
      
      MailApp.sendEmail(eMail, sub, resp);

      var _row = "A"+lastRow+":T"+lastRow
      sheet.getRange(lastRow, 22).setValue("System Failure");
      sheet.getRange(_row).setBackground('#ff5749'); // Light red
    }
    return;
  }
  //Utilities.sleep(0000);

  if (action == 1) { //Add
    evId = doCreateEvent(description, eMail, contactNum, title, startFirstDate, duration, room, recurring, freq, unit, occurencesDays, daysOfWeek, sameDateEveryMonth, lastRecurrenceDate);
    
    if (evId != null) {
      var sub = "Your Event Creation Request for Start Date : " + startFirstDate.toDateString() + " SUCCEEDED";
      var resp = "Location : Duet :" 
      if (room == 1)
       resp += "Big Room\n"
      else if (room == 2)
        resp += "Small Room\n"
      else 
        resp+= "Both Rooms\n"
    
      resp+= "Title : " + title + "\n"
      
      if (description)
        resp +="Description : " + description + " \n"
      resp += "Start Time : " + startFirstDate.toLocaleDateString() + " " + startFirstDate.toLocaleTimeString() + "\n"
      + "Duration : " + duration + "\n"
      + "Recurring :" + recurring + "\n" 
      + "Requester's Name : " + requester + "\n";
      if (fromGE_)
        resp += "From : " + fromGE_ + "\n";
      if (contactNum)
        resp+= "Contact Number : " + contactNum + "\n";

      var recDetails = "";
      if (recurring == 'Yes') {
        recDetails += "Repeats every ";
        if (unit == "Day(s)") {
          if (freq == 1)
            recDetails += "day ";
          else
            recDetails += freq + " days "
        }

        if (unit == "Week(s)") {
          if (freq == 1)
            recDetails += "Week ";
          else
            recDetails += freq + " Weeks "
          recDetails += "on ";
          _n = daysOfWeek.length
          if (_n == 1)
            recDetails+= daysOfWeek[0] + " ";
          else if (_n == 2)
            recDetails+= daysOfWeek[0] + " and " + daysOfWeek[1] + " ";
          else {
            for (_i = 0; _i < daysOfWeek.length; _i++) {
              recDetails+=  daysOfWeek[_i];
              if (_i == _n - 2)
                recDetails+= ", and ";
              else
                recDetails+= ", ";
            }
          }
        }
        if (unit == "Month(s)") {
          var _days = [
            CalendarApp.Weekday.MONDAY,
            CalendarApp.Weekday.TUESDAY,
            CalendarApp.Weekday.WEDNESDAY,
            CalendarApp.Weekday.THURSDAY,
            CalendarApp.Weekday.FRIDAY,
            CalendarApp.Weekday.SATURDAY,
            CalendarApp.Weekday.SUNDAY,
          ];
          if (freq == 1)
            recDetails += "Month ";
          else
            recDetails += freq + " Months "
          _getDate = startFirstDate.getDate();
          if (!sameDateEveryMonth) {
            _getDay = startFirstDate.getDay()
            idx = Math.floor((_getDate - 1)/7) + 1
            _suffix = ["st", "st", "nd", "rd", "th", "th"]
            if (idx > 4)
              recDetails += "on the last " + _days[_getDay ==0 ? 6 : _getDay - 1] + " "
            else
              recDetails += "on " + idx + " " + _suffix[idx] + " " + _days[_getDay ==0 ? 6 : _getDay - 1] + " "
          } else {
            _suffix = ["th", "st", "nd", "rd", "th", "th", "th", "th", "th", "th"]
            Logger.log("_getDate : " + _getDate)
            
            if (_getDate == 11)
              _sx = "th";
            else
              _sx = _suffix[_getDate%10];
            recDetails += "on " + _getDate + " " + _sx + " of every month "
          }
        }
        recDetails += "for a total of " + occurencesDays + " events "
        recDetails += "with the last event occuring on : " + lastRecurrenceDate.toLocaleDateString() + " at : " + lastRecurrenceDate.toLocaleTimeString() + ".\n"
        Logger.log(resp + recDetails)
      }
      
      recDetails += "\nIf you need to delete this event, please use this id " + evId[(room == 2)?1:0] + " to cancel the booking. Thank you."
      MailApp.sendEmail(eMail, sub, resp + recDetails)
      if (room == 1)
        sheet.getRange(lastRow, 21).setValue(evId[0]);
      else if (room == 2)
        sheet.getRange(lastRow, 22).setValue(evId[1]);
      else {
        sheet.getRange(lastRow, 21).setValue(evId[0]);
        sheet.getRange(lastRow, 22).setValue(evId[1]);
      }
    } else {
      var sub = "FAILED!! Your Event Creation Request for Start Date : " + startFirstDate.toDateString() + " failed.";
      
      var resp = "Location : Duet :" 
      var bigRoom = CalendarApp.getOwnedCalendarsByName('Big Room');
      var smallRoom = CalendarApp.getOwnedCalendarsByName('Small Room');
      var bURL =  "https://calendar.google.com/calendar/embed?src="+bigRoom[0].getId()+"&ctz=Asia%2FKolkata"
      var sURL= "https://calendar.google.com/calendar/embed?src="+smallRoom[0].getId()+"&ctz=Asia%2FKolkata"

      if (room == 1)
       resp += "Big Room\n"
      else if (room == 2)
       resp += "Small Room\n"
      else 
       resp+= "Both Rooms\n"


      resp+= "Title : " + title + "\n"
      if (description)
        resp +="Description : " + description + " \n"

      resp += "Start Time : " + startFirstDate.toLocaleDateString() + " " + startFirstDate.toLocaleTimeString() + "\n"
      + "Duration : " + duration + "\n"
      + "Recurring : " + recurring + "\n" 
      + "Requester's Name :" + requester + "\n"
      + "From : " + fromGE_ + "\n";
      if (contactNum)
        resp+= "Contact Number : " + contactNum + "\n";
      
      resp+="This likely happened because of a conflict. Please check the calendar for available date(s) and times. Thank you!\n"
            + "Calendar link : \n"
      if (room == 1)
        resp += "Big Room : " + bURL + "\n";
      else if (room == 2)
        resp+= "Small Room : " + sURL + "\n";
      else {
        resp += "Big Room : " + bURL + "\n";
        resp += "Small Room : " + sURL + "\n";
      }
      
      MailApp.sendEmail(eMail, sub, resp);
      Logger.log(resp)
      var _row = "A"+lastRow+":V"+lastRow
      
      sheet.getRange(_row).setBackground('#ff5749'); // Light red
    }
  } else { //Delete
    var bigRoom = CalendarApp.getOwnedCalendarsByName('Big Room');
    var smallRoom = CalendarApp.getOwnedCalendarsByName('Small Room');
    
    var textFinder = sheet.createTextFinder(delEventId).matchEntireCell(true);

    {
      var _range = textFinder.findAll();
      Logger.log(_range.length)
      var _r = -1;
      Logger.log("Deleting : " + delEventId);

      
      for (var _i = 0; _i < _range.length; _i++){
        _r = _range[_i].getRow();
        if (_range[_i].getColumn() == 13) {
          sheet.deleteRow(_r);
        }
        if (_range[_i].getColumn() == 22) {
          var _row = "A"+_r+":V"+_r;
          const _eventS = smallRoom[0].getEventSeriesById(delEventId);
          try {
            _eventS.deleteEventSeries();
          } catch(err) {
            Logger.log("Likely already deleted from the Small Room Cal!");
          }
          sheet.getRange(_row).setBackground('#c2c2c2'); //Light grey
        }

        if (_range[_i].getColumn() == 21) {
          var _row = "A"+_r+":V"+_r;
          const _eventB = bigRoom[0].getEventSeriesById(delEventId);
          try {
            _eventB.deleteEventSeries();
          } catch(err) {
            Logger.log("Likely already deleted from the Big Room Cal!");
          }
          if (sheet.getRange(_r,22).getValue()) {
            const _eventS = smallRoom[0].getEventSeriesById(sheet.getRange(_r,22).getValue());
            try {
              _eventS.deleteEventSeries();
            } catch(err) {
              Logger.log("Likely already deleted from the Small Room Cal!");
            }
          }
          sheet.getRange(_row).setBackground('#c2c2c2'); //Light grey
        }
      }
      lock.releaseLock();
      return;
    } 
  }
  lock.releaseLock();
  return;
}

function myBooking(e) {
  var rangeData = sheet.getDataRange();
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  if (!lock.hasLock()) {
    Logger.log("failed to get the lock after 30 seconds")
    return;
  }
  lastRow = rangeData.getLastRow();

  var lastColumn = rangeData.getLastColumn();

  if (lastColumn != 22) {
	  var newCol = sheet.insertColumnAfter(lastColumn);
	  sheet.getRange(1,21).setValue("Big Room Event Id");
    sheet.getRange(1,22).setValue("Small Room Event Id")
  }
  lock.releaseLock();
  parseFormResponse(lastRow);
}
