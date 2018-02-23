/*
https://gist.github.com/mhawksey/5705633
1. Set two variables below
2. Run > setup (twice, once to authenticate the script, second to actually run
3. Rescources > Current project trigger's and then click 'No triggers set up. Click here to add one now.'
You can accept the defaults to run emailAnnouncements() every hour
*/
 
 
var url_of_announcements_page = "https://sites.google.com/a/ahold.com/datawarehouse/home/news"; // where your site page is
var who_to_email = "scott.lawrence@aholdusa.com"; // who to send to (it can be a comma seperated list) AUSA.IDWMicrostrategy_Users.Group@ahold.com
//AUSA.SCDWMicrostrategy_Users.Group@ahold.com,AUSA.IDWMicrostrategy_Users.Group@ahold.com,AUSA.CDWMicrostrategy_Users.Group@ahold.com
 
function emailAnnouncements(){
  var page = SitesApp.getPageByUrl(url_of_announcements_page); // returns a page object
  if(page.getPageType() == SitesApp.PageType.ANNOUNCEMENTS_PAGE){ // test if page object is announcement page
    // get the last 10 announcements 
    var announcements = page.getAnnouncements({ start: 0,
                                               max: 10,
                                               includeDrafts: true, //true
                                               includeDeleted: false});
                                               
                                              
    announcements.reverse(); // reverse the result order so oldest first                                          
    for(var i in announcements) { // loop the individual announcements
      var ann = announcements[i]; // just creating a little shortcut
      var updated = ann.getLastUpdated().getTime(); // get when updated
      if (updated > PropertiesService.getScriptProperties().getProperty('last-update')){ // if greater than last update send email
        var options = {}; // options used in MailApp
        // create html body of the email using an announcement
        options.htmlBody = Utilities.formatString("<h1><a href='%s'>%s</a></h1><p>%s</p>", 'url', ann.getTitle(), ann.getHtmlContent());  //ann.getUrl()
        // send the email
        options.name = 'Data Warehouse Support';
        options.noReply = false;
        options.replyTo = '';
        options.bcc = '';
        MailApp.sendEmail(who_to_email, "MicroStrategy Announcement: "+ann.getTitle(), ann.getTextContent()+"\n\n"+ann.getUrl(), options);
        // update the last email sent
        PropertiesService.getScriptProperties().setProperty('last-update',updated);
      }
    }
  }
}
 
function setup(){
  PropertiesService.getScriptProperties().setProperty('last-update',new Date().getTime());
}
  
  
  
 

