/**
* Functions to get labels
* get emails
* convert to PDF
* send to Drive with attachments
* list in sheets
*
* @author Graeme A White
* email to PDF from https://ctrlq.org/code/19117-save-gmail-as-pdf
*
**/

function onOpen(e){
 var ui = SpreadsheetApp.getUi();
  ui.createMenu('GMail-to-PDF')
      .addItem('Open Side Bar...', 'setUpNewSidebar')
      .addToUi();
}

function setUpNewSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('newSidebar')
           .setTitle('Send Emails to Drive')
           .setSandboxMode(HtmlService.SandboxMode.IFRAME);
         SpreadsheetApp.getUi().showSidebar(ui);
} 

function getLabels() {

  var labels = GmailApp.getUserLabels();
  var arrayLabels = [];

if (labels.length < 0){ 
  Browser.msgBox('There are no Labels in Gmail! Create a label in Gmail add emails and try again');
      }
 
for (var i = 0; i < labels.length; i++) {
   arrayLabels.push ([labels[i].getName()]);
      }
return arrayLabels;
}

function getEmailDetails(gmailLabels) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listSheet = ss.getSheetByName("ListEmails");
  var settingsSheet = ss.getSheetByName("Settings");
  var emailSheet = ss.getSheetByName("emails");

  listSheet.getRange("A3:I1000").clear({contentsOnly: true});
  listSheet.getRange("B1:C1").clear({contentsOnly: true});
  emailSheet.getRange("A2:I1000").clear({contentsOnly: true});
  settingsSheet.getRange("A1:D1").clear({contentsOnly: true});  
  SpreadsheetApp.flush();
  
  var startTime = +new Date();
  settingsSheet.getRange("A1").setValue(startTime);
  SpreadsheetApp.flush();

  var emailArray = [];
  var emailSubArray = [];
  var threads = GmailApp.search("in:" + gmailLabels);
      
  if (threads.length > 0) {
   
    listSheet.getRange("B1").setValue(" PLEASE WAIT...CURRENTLY FETCHING " );
    SpreadsheetApp.flush();
    
       for (var t=0; t<threads.length; t++) {
       var msgs = threads[t].getMessages();
       var subject = threads[t].getFirstMessageSubject();
   
           for (var m=0; m<msgs.length; m++) {
              var msg = msgs[m];
               }
         
      var atts = msg.getAttachments();
      emailArray.push([t+1, msg.getSubject(),  msg.getFrom(),  msg.getTo(), msg.getDate(), atts.length, (new Date()), msg.getId()]); 
      
       var msgBody = msg.getBody().replace(/<.*?>/g, '\n')
                .replace(/^\s*\n/gm, '').replace(/^\s*/gm, '').replace(/\s*\n/gm, '\n');
      
      emailSubArray.push([t+1, msg.getFrom(), msg.getTo(), msg.getCc(), msg.getDate(), msg.getSubject(),  msgBody ]);
      
          }
        listSheet.getRange(3, 1, emailArray.length, 8 ).setValues(emailArray);   
        emailSheet.getRange(2,1, emailSubArray.length, 7).setValues(emailSubArray); 
         
        //call function to write folder to drive       
        createFolder(gmailLabels);
       
      }
 else{
   Browser.msgBox('This Label may not exist or there are no emails in Gmail Label: ' 
   +gmailLabels +'.  Please make sure you selected the correct label from the'  
   +' list above and added emails to this label and try again');
   }
}

function createFolder(gmailLabels){

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName("Settings");
 
    var fileDate = Utilities.formatDate(new Date(), "GMT+11", "dd/MM/yyyy' 'HH:mm  aa");
    var driveFolder  = "From Gmail - Label: " +gmailLabels +" Accessed: " +fileDate;
    var folders = DriveApp.getFoldersByName(driveFolder);
    
    var folder = folders.hasNext() ? 
        folders.next() : DriveApp.createFolder(driveFolder);
        
        // need to set gmailLables and drive folder to settings
        
    settingsSheet.getRange('B1').setValue(gmailLabels);
    settingsSheet.getRange('C1').setValue(driveFolder);
    SpreadsheetApp.flush();
  
  //Create trigger to run get Emails every ten Minutes if required with a large number of emails
    createTrigger_("getEmails",10);
  // get emails and write to PDF
    getEmails();
    
    }
    
function getEmails() {
    
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var listSheet = ss.getSheetByName("ListEmails");
   var settingsSheet = ss.getSheetByName("Settings");
  
   var lastRow = listSheet.getLastRow();
   var data = listSheet.getRange(3, 1, (lastRow-2), 9).getValues();
   var msgs = [];
   var gmailLabels = settingsSheet.getRange('B1').getValue();
   var driveFolder = settingsSheet.getRange('C1').getValue();
   
    // Prefix '+' to get date as epochy number.
  var elapsedTime, startTime = +new Date();

   // Google Drive folder where the Files are to be saved 
   var folderUrl = DriveApp.getFoldersByName(driveFolder).next().getUrl(); 
   var folder = DriveApp.getFoldersByName(driveFolder).next();
        
   // Append all  messages in emails sheet / write to drive
   for (var m=0; m<data.length; m++) {
   
    var written = data[m][8];
    
     if (written !== "YES") {
   
      var msg = GmailApp.getMessageById(data[m][7]);
      var subject = (data[m][1]);
      var html = "" ;
      var attachments = [];
         
        html += "From: " + msg.getFrom().replace(/<|>/g, "'") + "<br />";  
        html += "To: " + msg.getTo().replace(/<|>/g, "'") + "<br />";
        html += "cc: " + msg.getCc().replace(/<|>/g, "'") + "<br />";
        html += "bcc: " + msg.getBcc().replace(/<|>/g, "'") + "<br />";
        html += "Date: " + msg.getDate() + "<br />";
        html += "Subject: " + msg.getSubject() + "<br />"; 
        html += "<hr />";
        html += msg.getBody().replace(/<img[^>]*>/g,"");
        html += "<hr />";
        
        var atts = msg.getAttachments();
        for (var a=0; a<atts.length; a++) {
             attachments.push(atts[a]);
             }
      
      // Save the attachment files and create links in the document's footer 
      if (attachments.length > 0) {
         var footer = "<strong>Attachments:</strong><ul>";
       
         var subFolderName = "Attachments: " + "(" +(m+1) +")" +subject;
         var subFolder = folder.createFolder(subFolderName);
         var subFolders = DriveApp.getFoldersByName(subFolderName).next();
       
          for (var z=0; z<attachments.length; z++) {
          
             var file = subFolders.createFile(attachments[z]);
             footer += "<li><a> " + file.getName() + "</a></li>"
             }
          html += footer + "</ul>";
        }
      
      // Convert the Email Thread into a PDF File 
      var tempFile = DriveApp.createFile("temp.html", html, "text/html");
     folder.createFile(tempFile.getAs("application/pdf")).setName("(" +(m+1) +")" +subject + ".pdf");
     tempFile.setTrashed(true); 
    
     listSheet.getRange((m+3), 9).setValue("YES");
     listSheet.getRange("B1").setValue(" PLEASE WAIT...CURRENTLY FETCHING " + (m+1) + " OUT OF " + (data.length));
     SpreadsheetApp.flush();
     }
            
       // Recalculate elapsedTime.
    elapsedTime = +new Date() - startTime;
    if(elapsedTime> 300000){ // 300000 ms or 5 minutes.
      return;
      }    
    }
    
      //list the drive file and link
     listSheet.getRange("B1").setValue(driveFolder);
     listSheet.getRange("C1").setValue(folderUrl);
     //Get an end time and place in the settings sheet
     var endTime = +new Date();
     settingsSheet.getRange("D1").setValue(endTime);
     SpreadsheetApp.flush();
     
      // Loop completed successfully, so delete trigger.
     deleteTriggers_();   
}

function deleteTriggers_(){
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger){
    ScriptApp.deleteTrigger(trigger);
    Utilities.sleep(1000); // In millisecond.
  });
 }

 function createTrigger_(funcName,minutes){
// Delete already created triggers if any.
 deleteTriggers_();
   ScriptApp.newTrigger(funcName).timeBased()
    .everyMinutes(minutes).create();
 }