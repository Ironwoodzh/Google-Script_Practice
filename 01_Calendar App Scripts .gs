function onOpen(e) {
  SpreadsheetApp.getUi().createMenu("CalendarSetup")
      .addItem('GetYourCalendarEvents', 'getCalEvents')
      .addSeparator()
      .addItem('AddYourNewEventsToCalendar', 'addCalEvents')
      .addSeparator()
      .addItem('AddNewGuest&SendEmailToThemOnly', 'addGuestOnly')
      .addSeparator()
      .addItem('DeleteYourCalendarEvents', 'deleteCalEvents')
      .addSeparator()
      .addItem('EditYourCalendarEvents', 'editCalEvents')
      .addSeparator()
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function getCalEvents() {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GetCalendar");
  let calid = ss.getRange(4,2).getValues()
  let cal = CalendarApp.getCalendarById(calid)


  // input Date from the sheet 
  let setsd = new Date(ss.getRange(2,2).getValues())
  let seted = new Date(ss.getRange(3,2).getValues())
  let events =cal.getEvents(setsd,seted);

  // Clear content before get event to sheet
  let lr = ss.getLastRow();
  let lc = ss.getLastColumn();
  ss.getRange(6, 1, lr,lc).clearContent();

  // get events and input to the sheet
  for(let i = 0;i<events.length;i++){
    
    let wd = events[i].getStartTime();
    let sd = events[i].getStartTime();
    let ed = events[i].getEndTime();
    let dur = (ed-sd)/1000/60/60

    let title = events[i].getTitle();

    let creator = events[i].getCreators();

    let guests = events[i].getGuestList();
    let guestEmails = "";
    for (let j = 0; j < guests.length; j++) {
        let guest = guests[j].getEmail();
        guestEmails += guest+", ";
    }

    let loc = events[i].getLocation();
    let des = events[i].getDescription();
    let eventid = events[i].getId();
   

    ss.getRange(i+6,1).setValue(eventid.replace("@google.com",""));
    ss.getRange(i+6,2).setValue(sd).setNumberFormat("yyyy/mm/dd DDD hh:mm:ss");
    ss.getRange(i+6,3).setValue(ed).setNumberFormat("yyyy/mm/dd DDD hh:mm:ss");
    ss.getRange(i+6,4).setValue(dur);
    ss.getRange(i+6,5).setValue(Utilities.formatDate(new Date(wd), "GMT", "w")-1);
    ss.getRange(i+6,6).setValue(title);
    ss.getRange(i+6,7).setValue(creator);
    ss.getRange(i+6,8).setValue(guestEmails);
    ss.getRange(i+6,9).setValue(loc);
    ss.getRange(i+6,10).setValue(des);

  }

}

function addCalEvents() {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AddCalendar");
  let lr = ss.getLastRow();
  let calid = ss.getRange(2,2).getValues()

  let cal = CalendarApp.getCalendarById(calid)

  let data = ss.getRange("A4:G"+ lr).getValues();
  for(let i = 0;i<data.length;i++){
    cal.createEvent(data[i][0],data[i][1],data[i][2], {location: data[i][3], description: data[i][4], guests: data[i][5], sendInvites: data[i][6]})
  }
}

function addGuestOnly(){

  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AddGuestOnly");
  let calid = ss.getRange(4,2).getValues()
  let cal = CalendarApp.getCalendarById(calid)
  let lr = ss.getLastRow();

  // input Date from the sheet 
  let setsd = new Date(ss.getRange(2,2).getValues())
  let seted = new Date(ss.getRange(3,2).getValues())
  let events =cal.getEvents(setsd,seted);

  let sheetIds = ss.getRange(6,1,lr).getValues()

  // get events
  for(let i = 0;i<events.length;i++){
    let eventid = events[i].getId();

    for(k=0; k<sheetIds.length;k++){
      let sheetid = sheetIds[k] 
      let sheetAddGuest = ss.getRange(6+k,11).getValue()

      if(sheetid+"@google.com" == eventid){
        
        if(sheetAddGuest != ""){
        Logger.log("Guest Add:"+sheetAddGuest)
        addGuestAndSendEmail(calid,sheetid,sheetAddGuest)
        }
      }
    }
  }
}

function deleteCalEvents(){
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DeleteCalendar");
  let lr = ss.getLastRow();

  // setup calendar by id
  let calid = ss.getRange(4,2).getValues()
  let cal = CalendarApp.getCalendarById(calid)

  // input Date from the sheet 
  let fromDate = new Date(ss.getRange(2,2).getValues())
  let toDate = new Date(ss.getRange(3,2).getValues())
  let events = cal.getEvents(fromDate,toDate);
  
  let sheetIds = ss.getRange(6,1,lr).getValues()
  let sheetEventTitle = ss.getRange(6,2,lr).getValues();
  
  ss.getRange(6, 3, lr).clearContent();
  

  for(let i=0; i<sheetIds.length; i++){
    let sheetId = sheetIds[i]
    Logger.log(sheetIds)
    for (let j = 0; j < events.length; j++) {
      let event = events[j];
      Logger.log(event.getId()+[j])
      if(sheetId + "@google.com" == event.getId()){
        Logger.log("Delete: "+ event.getTitle()+"_"+[j+1])
        ss.getRange(j+6,3).setValue("Delete: "+[j+1]+"_"+ event.getTitle()+event.getStartTime());
        event.deleteEvent()
      }
    }
  }

}



function editCalEvents(){
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EditCalendar");
  let calid = ss.getRange(4,2).getValues()
  let cal = CalendarApp.getCalendarById(calid)
  let lr = ss.getLastRow();

  // input Date from the sheet 
  let setsd = new Date(ss.getRange(2,2).getValues())
  let seted = new Date(ss.getRange(3,2).getValues())
  let events =cal.getEvents(setsd,seted);

  let sheetIds = ss.getRange(6,1,lr).getValues()
  let sheetSds = ss.getRange(6,2,lr).getValues()
  let sheetEds = ss.getRange(6,3,lr).getValues()

  // get events
  for(let i = 0;i<events.length;i++){
    let eventid = events[i].getId();

    for(k=0; k<sheetIds.length;k++){
      let sheetid = sheetIds[k] 
      let sheetSd = sheetSds[k]
      let sheetEd = sheetEds[k]
      let sheetEventTitle = ss.getRange(6+k,6).getValue()
      let sheetCreator = ss.getRange(6+k,7).getValue()
      let sheetEventGuest = ss.getRange(6+k,8).getValue()
      let sheetLocation = ss.getRange(6+k,9).getValue()
      let sheetDescription = ss.getRange(6+k,10).getValue()
      let sendUpdates = ss.getRange(6+k,11).getValue()
      let visibility = ss.getRange(6+k,12).getValue()
      let colorID = ss.getRange(6+k,13).getValue()

      if(sheetid+"@google.com" == eventid){
        
        updateAndSendEmail(
        calid,
        sheetid,
        sheetSd,
        sheetEd,
        sheetEventTitle,
        sheetCreator,
        sheetEventGuest,
        sheetLocation,
        sheetDescription,
        sendUpdates,
        visibility,
        colorID
        )
      
      }
    }
  }
}


// // Note: requires advanced API
function updateAndSendEmail(
  calendarId,
  eventId,
  sheetSd,
  sheetEd,
  sheetEventTitle,
  sheetCreator,
  sheetEventGuest,
  sheetLocation,
  sheetDescription,
  sendUpdates,
  visibility,
  colorId
  ) {
  var calendarId = calendarId.toString()
  var sheetEventGuest = sheetEventGuest.split(",")
  var event = Calendar.Events.get(calendarId, eventId);
  var attendees = [event.attendees];
  
  if(sheetEventGuest!=""){
    for(let i=0;i<sheetEventGuest.length;i++){
      var ng = sheetEventGuest[i]
      attendees.push({email: ng})
    }
  }else {
    attendees = ""
  }
  if(visibility != true){
    visibility ="public"
    } else{
    visibility ="private"
    }
  
  var colorId = colorId.replace(/[^0-9]/ig,"")
  var resource =  {
    start: {dateTime: (new Date(sheetSd)).toISOString()},
    end: {dateTime: (new Date(sheetEd)).toISOString()},
    summary: sheetEventTitle,
    creator: sheetCreator,
    location: sheetLocation,
    description: sheetDescription,
    attendees: attendees,
    visibility: visibility,
    colorId: colorId
  };
    if(sendUpdates != true){
      sendUpdates ="none"
    } else{
      sendUpdates ="all"
    }
  var args =  {sendUpdates: sendUpdates}

  Calendar.Events.update(resource,calendarId, eventId,args);
}

// Note: requires advanced API
function addGuestAndSendEmail(calendarId, eventId, newGuest) {
  var calendarId = calendarId.toString()
  var newGuest = newGuest.split(",")

  var event = Calendar.Events.get(calendarId, eventId);
  var attendees = [event.attendees];
  for(let i=0;i<newGuest.length;i++){
    var ng = newGuest[i]
    attendees.push({email: ng})
  }
  
  var resource = {attendees: attendees};
  var args = { sendUpdates: "all" };

  Calendar.Events.patch(resource, calendarId, eventId, args);
}
