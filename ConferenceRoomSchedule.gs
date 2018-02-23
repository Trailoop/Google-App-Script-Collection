/*************************************************************************************************************************************
                                                  Conference Room Daily Schedule EMAIL APP
                                                         Google App Script (GAS)
                                                        Created by Scott Lawrence
                                             Subject Matter Expert - Oracle Project Finance
                                                   Contact: scott.lawrence@aholdusa.com
 
                                                   Scheduling GAS - time trigger setup  
                                                     Function - sendConferenceRoomSchedule
                                 Timing: Weekday time triggers run the function Monday- Friday at 6am - 7am.
                                 
                                 
                                 http://docs.oracle.com/javase/6/docs/api/java/text/SimpleDateFormat.html
                                 https://developers.google.com/apps-script/class_utilities#formatDate
 
 ======================================================================================================================================*/
function tz () { 

Logger.log(Session.getTimeZone());

var d = new Date()
var n = d.getTimezoneOffset();

Logger.log(n);

}

function findCalendars () {
  var calendars = CalendarApp.getAllCalendars();
  
  for(var i = 0; i < calendars.length;i++) {
    Logger.log(calendars[i].getName() + ":" + calendars[i].getId() + "\n");
  }
}

//Class ConferenceRoom
function ConferenceRoom (name,id) {
  this.name = name;
  this.id = id;
  this.getTodaysEvents = function (day){
    var room = CalendarApp.getCalendarById(this.id);
    var events = room.getEventsForDay(day);
    
    return events;
  };
  this.getHtmlTodayEvents = function (day) {
    var events = this.getTodaysEvents(day);
    var htmlEvents = "<p><h4>" + this.name + "</h4><ul>";
    
    for (var e in events) {
        htmlEvents += "<li>" + events[e].getTitle() +"<strong> : </strong>" + Utilities.formatDate(events[e].getStartTime(),"America/New_York","h:mm a");
        htmlEvents += " - " + Utilities.formatDate(events[e].getEndTime(),"America/New_York","h:mm a") + "</li>"; 
      }
    
    htmlEvents += "</ul></p>";
    
    return htmlEvents;
  
  };
  this.getHtmlName = function () {
    return "<th>" + this.name + "</th>";
  };
}

        
function getConferenceRoomsEvents () {
  var today = new Date();
  var cantonSecondFloorRoomA = new ConferenceRoom("Canton 2nd Floor A","ahold.com_2d32383236383831342d3732@resource.calendar.google.com");
  var cantonSecondFloorRoomAEvents = cantonSecondFloorRoomA.getHtmlTodayEvents(today);
    
  var cantonSecondFloorRoomB = new ConferenceRoom("Canton 2nd Floor B","ahold.com_2d37363134373039373130@resource.calendar.google.com");
  var cantonSecondFloorRoomBEvents = cantonSecondFloorRoomB.getHtmlTodayEvents(today);
  
  var cantonFirstFloorRoomA = new ConferenceRoom("Canton 1st Floor A","ahold.com_2d37313734383636362d373036@resource.calendar.google.com");
  var cantonFirstFloorRoomAEvents = cantonFirstFloorRoomA.getHtmlTodayEvents(today);
  
  var cantonFirstFloorRoomB = new ConferenceRoom("Canton 1st Floor B","ahold.com_2d32343838343232352d343634@resource.calendar.google.com");
  var cantonFirstFloorRoomBEvents = cantonFirstFloorRoomB.getHtmlTodayEvents(today);
  
  var cantonFirstFloorRoomC = new ConferenceRoom("Canton 1st Floor C","ahold.com_35353033333935392d323133@resource.calendar.google.com");
  var cantonFirstFloorRoomCEvents = cantonFirstFloorRoomC.getHtmlTodayEvents(today);
  
  var cantonFirstFloorRoomE = new ConferenceRoom("Canton 1st Floor E","ahold.com_3534373332363839@resource.calendar.google.com");
  var cantonFirstFloorRoomEEvents = cantonFirstFloorRoomE.getHtmlTodayEvents(today);
  
  var cantonFirstFloorRoomLarge = new ConferenceRoom("Canton 1st Floor Large","ahold.com_2d3231313835343838353936@resource.calendar.google.com");
  var cantonFirstFloorRoomLargeEvents = cantonFirstFloorRoomLarge.getHtmlTodayEvents(today);

  var html  = cantonSecondFloorRoomAEvents + cantonSecondFloorRoomBEvents + cantonFirstFloorRoomAEvents + cantonFirstFloorRoomBEvents;
      html += cantonFirstFloorRoomCEvents + cantonFirstFloorRoomEEvents + cantonFirstFloorRoomLargeEvents;
  
  return html;
}

function sendConferenceRoomSchedule () {
  var todaydate = new Date();
  var day = ((todaydate.getDate()<10) ? "0" : "")+ todaydate.getDate();
  var month = (((todaydate.getMonth() + 1) <10) ? "0" : "") + (todaydate.getMonth()+ 1);
  var year = todaydate.getYear ();
  var datestamp = month + "." + day + "." + year;
  var recipients = "";
  var bcc = "scott.lawrence@aholdusa.com,lchruney@aholdusa.com,kathleen.connors@ahold.com,mheffern@ahold.com,jwhitten@aholdusa.com";
  
  var header = (<r><![CDATA[ 
    <html>
    <head>
    </head>
    <body>
    <table style="font-family:Helvetica;-webkit-border-radius: 5px;border-radius: 5px;">
      <thead style="-webkit-border-radius: 5px;border-radius: 5px;text-align:center">
      </thead>
      <tbody style="font-family: “Lucida Grande”, sans-serif;font-size: 11.67px;font-style: normal;font-weight: normal;text-transform: normal;letter-spacing: normal;line-height: 1.4em;">
      
        
  ]]></r>).toString();
  
  var footer = (<r><![CDATA[ </tbody></table></body></html>]]></r>).toString();
  
  //email contents - non-html 
  var message = "This is an automatically generated email, please do not reply.\r\n\r\n" +
      "Please find the attached Master Action Item Log for " + datestamp + "\r\n\r\n\r\n\r\n" + 
      "If you would like to be removed from the distribution send a message to Scott Lawrence, scott.lawrence@aholdusa.com\r\n";
    
  var htmlmessage = header + getConferenceRoomsEvents() + footer;                        
  
  MailApp.sendEmail(recipients, "Oracle Project: Conference Room daily Schedule " + datestamp, message, {name: "Oracle Project Finance", bcc: bcc, noReply: true, htmlBody: htmlmessage});
}



