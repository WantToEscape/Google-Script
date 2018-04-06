
//  Explain : pdf Data convert to Text and send to email
//  Author : Arthur
//  Date : 01-30-2018
//  Must have gmail box; 
//  Must have google drive and exist folder [01 - PDF Raw Data],folder name can be optional,Create new folder in google drive and reset name "01 - PDF Raw Data";
//  PDF content mustn't be image

function pdfConvertText() {
  var pdfFolderName="01 - PDF Raw Data";  //  Save PDF files in this forlder on google drive that this folder must be exist
  var txtFolderName="02 - Text Data";
  
  //  Delete existed txt folder
  var folderIter=DriveApp.getFoldersByName(txtFolderName);
  while (folderIter.hasNext()) {
   folderIter.next().setTrashed(true);
  }

  
  var folderIter=DriveApp.getFoldersByName(pdfFolderName);
  var folder=folderIter.next();
  var filesIter=folder.getFiles();
  var count=0;
  
  //  Loop convert all files
  while(filesIter.hasNext()){
    var file=filesIter.next();
    var filename=file.getName();
    
    if(file.getMimeType()=="application/pdf"){
      count+=1
      
      //  var myID=file.getId();
      //  Logger.log(myID)
      //  Use Drive API to convert
      var resource = {
        title: filename,
        mimeType: "application/pdf"
      };
      var file = Drive.Files.insert(resource, file, {ocr: true, ocrLanguage: "en"});
      
      //  Extract Text from PDF file
      var doc = DocumentApp.openById(file.id);
      var text = doc.getBody().getText();
      
      //  Remove redundant files 移除冗余文件
      var idToDle = file.getId();
      var removeID = Drive.Files.remove(idToDle);
      
      
      //  Save PDF TEXT to txt
      var myFolderName ="02 - Text Data"
      filename += ".txt"
      saveAsTXT(myFolderName,filename,text)
    }
  }
  
  if(count>0){
    autoSendEmail(txtFolderName,count);
  }
  //  Logger.log(count)
  
}



function saveData(folder, filename, contents) {
  var children = folder.getFilesByName(filename);
  var file = null;
  if (children.hasNext()) {
    file = children.next();
    file.setContent(contents);
  } else {
    file = folder.createFile(filename, contents);
  }
}


function saveAsTXT(folderName,fileName,contents) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    var folder = folders.next();
    saveData(folder, fileName, contents);
  } else {
    //  递归,不存在文件夹则创建
    var folders=DriveApp.createFolder(folderName)
    saveAsTXT(folderName,fileName,contents)
  }
}


//  Automatic send attachment of text/plain to gmail  
function autoSendEmail(txtFolderName,count){
  var myEmail = Session.getActiveUser().getEmail();
  var folderIter=DriveApp.getFoldersByName(txtFolderName);
  var folder=folderIter.next();
  var filesIter=folder.getFiles();
  var attachements = [];
  while(filesIter.hasNext()){
    var file=filesIter.next();
    //  var filename=file.getName();
    //  var myID=file.getId();
    //  var myFile = DriveApp.getFileById(myID);
    attachements.push(file.getAs("text/plain"));
  }
  
  //  SendEmail Parameter(Email,Subject,Content,{Attachment and the other infor})  参数(邮箱地址,标题,内容,附件及其它信息)
  MailApp.sendEmail(myEmail, 'Attachment Text For Reference', "Total " + count + " PDF files has been converted,there is automatically sent by the script function, you can ignore this information if you needn't run the macro.", {
     name: 'Automatic Emailer Script',    //  The name of the sender by default display
     attachments: attachements
  });
}


//function extractTextFromPDF() {
//   
//   //  PDF File URL 
//   //  You can also pull PDFs from Google Drive
//   var url = "https://img.labnol.org/files/Most-Useful-Websites.pdf";  
//   
//   var blob = UrlFetchApp.fetch(url).getBlob();
//   var resource = {
//     title: blob.getName(),
//     mimeType: blob.getContentType()
//   };
//   
//   //  Enable the Advanced Drive API Service
//   var file = Drive.Files.insert(resource, blob, {ocr: true, ocrLanguage: "en"});
//   
//   //  Extract Text from PDF file
//   var doc = DocumentApp.openById(file.id);
//   var text = doc.getBody().getText();
//   
//   return text;
//}
