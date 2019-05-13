var activeTestsCount = 0;
var runnedTestCount = 0;

var logKey = "";

function initializeGlobalVariables(){
  if(logKey == ""){
  
    var d = new Date();
    logKey = d.getHours() + "_" + d.getMinutes() + "_" + d.getSeconds() + "_" + d.getMilliseconds();
  }

  if(activeTestsCount == 0){

    for(var i = 0; i < Project.TestItems.ItemCount; i++){
      if(Project.TestItems.TestItem(i).Enabled){
        activeTestsCount++;
      }
    }
  }
}

function GeneralEvents_OnStopTest(Sender) 
{

  initializeGlobalVariables();

  runnedTestCount ++;


  Log.Message("Active tests count is " + activeTestsCount );
  Log.Message("Running test case " + runnedTestCount + " - " + Project.TestItems.Current.Name);

  if( activeTestsCount == runnedTestCount){

  Log.Message(Project.Path)

  var projectFileNameArray = Project.FileName.split(/(\\|\/)/g);  
  
  var productName = projectFileNameArray[projectFileNameArray.length -1].split('.')[0];
  
  
  var logPath = GetLogPath(productName,logKey);
  
  Log.Message("Log path  " + logPath );

  var saveResult = PackResults(logPath);

  var savedLogResult = GetLogItems(logPath);
  
  

  Log.Message(logPath + logKey +"\\site.mht");
  if (SendEmail("ditec.swd422@gmail.com", 
                        "ditec.swd422@gmail.com", 
                        "Test Results Notification-Project " + productName, 
                        "<html>Hello QA, here is result:" + "<br/><br/>" + savedLogResult + saveResult + "</html>",
                        logPath + logKey + "\\site.mht"))
    Log.Message("Mail was sent");
  else
    Log.Warning("Mail was not sent");

  }
}

function SendEmail(mFrom,  mTo,  mSubject,  mBody,  mAttachment)
{
  var smtpServer, portNumber, userName, userPassword;
  var useAutentification, useSSL, connectionTimeout;
  var i, schema, mConfig, mMessage;
  
  smtpServer = "smtp.gmail.com";
  smtpPort = 465;
  userLogin = "ditec.swd422@gmail.com"; // e.g. abc@gmail.com
  userPassword = "Test1234+";
  autentificationType = 1; // cdoBasic
  connectionTimeout = 30;
  // Required by GMail
  useSSL = true;

  try {
    schema = "http://schemas.microsoft.com/cdo/configuration/";
    mConfig = Sys.OleObject("CDO.Configuration");
    mConfig.Fields.Item(schema + "sendusing") = 2; // cdoSendUsingPort
    mConfig.Fields.Item(schema + "smtpserver") = smtpServer;
    mConfig.Fields.Item(schema + "smtpserverport") = smtpPort;
    mConfig.Fields.Item(schema + "sendusername") = userLogin;
    mConfig.Fields.Item(schema + "sendpassword") = userPassword;
    mConfig.Fields.Item(schema + "smtpauthenticate") = autentificationType;
    mConfig.Fields.Item(schema + "smtpusessl") = useSSL;
    mConfig.Fields.Item(schema + "smtpconnectiontimeout") = connectionTimeout;
    mConfig.Fields.Update();
   
    mMessage = Sys.OleObject("CDO.Message");
    mMessage.Configuration = mConfig;
    mMessage.From = mFrom;
    mMessage.To = mTo;
    mMessage.Subject = mSubject;
    mMessage.HTMLBody = mBody;    

    aqString.ListSeparator = ",";
    for(i = 0; i < aqString.GetListLength(mAttachment); i++)
      mMessage.AddAttachment(aqString.GetListItem(mAttachment, i));
    
    if(0 < mAttachment.length) {
      mMessage.AddAttachment(mAttachment);
    }
  
    mMessage.Send();
    
   }
   catch(exception) { 
      Log.Error("E-mail cannot be sent, " + "popis chyby: " +  exception.description);
     return false;  
   }
   Log.Message("Message to <" + mTo +  "> was successfully sent");
   return true;
}

function GetLogPath(productName, testKey)
{
  var WorkDir, DateDir, FileList, FileName, ArchivePath;
   var cDate = new Date();
   
   DateDir = String(cDate.getFullYear());
   DateDir += String((cDate.getMonth() < 9 ? '0' : '') + String((cDate.getMonth() + 1))); 
   DateDir += String((cDate.getDate() < 10 ? '0' : '') + String(cDate.getDate()));

   //vytvaranie suborov
   var fso  = new ActiveXObject("Scripting.FileSystemObject");
   if(!fso.FolderExists(Project.Path + "\\TestArchiveResults")){
     fso.CreateFolder(Project.Path + "\\TestArchiveResults");
   }
   if(!fso.FolderExists(Project.Path + "\\TestArchiveResults\\" + DateDir)){
     fso.CreateFolder(Project.Path + "\\TestArchiveResults\\" + DateDir);
   }
   fso.CreateFolder(Project.Path + "\\TestArchiveResults\\" + DateDir + "\\" + logKey);  
   
   
   WorkDir = "TestArchiveResults\\";  
//   WorkDir = productName + "/";  
   //return Project.Path + WorkDir + DateDir + "\\" + testKey + "\\" ;  
   return Project.Path + WorkDir + DateDir + "\\";  
}

function PackResults(logPath)
{
   
   //var TestResultsLink = "<br/>Final result &nbsp;&nbsp;<br/><a href=" + logPath + "index.htm\">Click here to open</a>";
   //var tmp2 = logPath.substr(20);
   //var tmp = logPath.replace(/\\/g,"/");
   var tmp = logPath;
   var TestResultsLink = "<br/>Final result &nbsp;&nbsp;<br/><a href=\"http://" + tmp + logKey + "index.htm\">Click here to open</a>";
   Log.Message("http:///" + tmp + logKey + "\\index.htm")
   Log.Message(tmp)
   //var TestResultsLink = "<br/>Final result &nbsp;&nbsp;<br/><a href=\"http://192.168.10.150/" + tmp + "index.htm\">Click here to open</a>";
   
   Log.SaveResultsAs(logPath + logKey, lsHTML);
   Log.SaveResultsAs(logPath + logKey + "\\site.mht", lsMHT);

   Log.SaveResultsAs(logPath + logKey + "\\root.xml", lsXML,true,lesCurrentTestItem);   
   
   //compress file
   var WorkDir = logPath + logKey;
  
   var FileList = slPacker.GetFileListFromFolder(WorkDir);
   var ArchivePath = WorkDir + "PackResults";
   if (slPacker.Pack(FileList, WorkDir, ArchivePath))
      Log.Message("The files have been compressed successfully");

return TestResultsLink;
}
 
function GetLogItems(logPath)
{
  /*
  var tempFolder = aqEnvironment.GetEnvironmentVariable("temp") + "\\" +   GetTickCount() + "\\";
                    

  
  if (0 != aqFileSystem.CreateFolder(tempFolder)) {
    Log.Error("The " + tempFolder + " temp folder was not created");
    return "";
  }
  if (!Log.SaveResultsAs(tempFolder, lsHTML)) {
    Log.Error("Log was not exported to the " + tempFolder + " temp folder");
    return "";
  }
  */

   
  var xDoc = Sys.OleObject("Msxml2.DOMDocument");
  
  //var xDoc = new ActiveXObject("msxml2.DOMDocument.4.0");
  
  
  Log.Message("ROOT XML : " + logPath + logKey + "\\root.xml");
  xDoc.load(logPath + logKey + "\\root.xml");                                            
  var result = LogDataToText(xDoc.childNodes.item(1), 0, "  ");
    
  //aqFileSystem.DeleteFolder(tempFolder, true);
   
  return "<table style=\"border-collapse:collapse;\"><tr><th style=\"border:#000 1px solid; text-align:left;background:#01b0f1;padding:4px\">Name</th><th style=\"border:#000 1px solid; text-align:left;background:#01b0f1;padding:4px\">Browser</th><th style=\"border:#000 1px solid; text-align:left;background:#01b0f1;padding:4px\">Status</th></tr>" + result + "</table>";   
}
 
function LogDataToText(logData, indentIndex, indentSymbol)
{

  if ("LogData" != logData.nodeName) {
    return "";
  }
 
  var result = "";
  
  
  for(var i = 0; i < indentIndex; i++) {
    result += indentSymbol;
  }
  var browsertmp = logData.getAttribute("name").search("Chrome");
  if ( browsertmp > 0 ){
      var browser = "chrome";
  }else {
    var browsertmp = logData.getAttribute("name").search("Firefox");
    if ( browsertmp > 0 ){
      var browser = "firefox";
    }else {
        var browsertmp = logData.getAttribute("name").search("IE");
        if ( browsertmp > 0 ){
          var browser = "iexplorer";
//          var browser = logData.getAttribute("name").substr(browsertmp);
        }else {
          var browsertmp = logData.getAttribute("name").search("Safari");
          if ( browsertmp > 0 ){
            var browser = "safari";
          }else { 
            var browser = "";
          }
        } 
    }
  }
  var line = logData.getAttribute("name").search("Keyword Test Log");
  if ( line != 0 ) 
  result = result + "<tr><td style=\"border:#000 1px solid;padding:4px\">" + logData.getAttribute("name") + "</td>" +
          "<td style=\"border:#000 1px solid; padding: 4px\">" + browser + "</td>" + 
            GetTextOfStatus(logData.getAttribute("status")) + "</tr>";
  
  for(var i = 0; i < logData.childNodes.length; i++) {
    result += LogDataToText(logData.childNodes.item(i), indentIndex + 1, 
                              indentSymbol);
  }
  return result;
}
 
function GetTextOfStatus(statusIndex) 
{
  switch(statusIndex) {
    case "0": return "<td style=\"background:#91d24d; border:#000 1px solid;padding:4px;text-align:center\">OK</td>";
    case "1": return "<td style=\"background:yellow; border:#000 1px solid;padding:4px;text-align:center\">WARNING</td>";
    case "2": return "<td style=\"background:#f00; border:#000 1px solid;padding:4px;text-align:center\">FAILED</td>";
    default: return "<td style=\"background:#00f; border:#000 1px solid;padding:4px;text-align:center\">UNDEFINED</td>";
  }
}



