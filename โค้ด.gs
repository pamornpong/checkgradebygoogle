//สร้าง PDF
function runPDF(){
 let sheetId = '1xLoUgJbiWS8ytyhis93X8lGH8jV3VvIo4WqzM9p0Fgc'
// ไอดีชองชีท จากข้อ 1
 let tmpFileId = '1Mg0Cqb5foUrRA0Xcrt59qWIkSHcMbZzeEhv5jqPinVs'
 //ไอดีของ slide จากข้อ 2
 let pdfFolder = DriveApp.getFoldersByName('ไฟล์ผลการเรียน1-64').next()
 let templateFile = DriveApp.getFileById(tmpFileId)
 //let data =PdfService.initData(sheetId, 'm31') 
 let data =PdfService.initData(sheetId, 'P6', 2, 6)
//ไม่ระบุเลขแถว เพราะต้องการสร้างไฟล์ PDF แค่แถวสุดท้าย
 let option = {
  pdfFolder: pdfFolder,
  templateFile: templateFile,
  data: data,
 // image_column: ['รูปภาพประจำตัว'],
  fileName: ['ผลการเรียน ','ชื่อ สกุล']
}
  PdfService.createPDFFromSlide(option)
}

function doGet(e) {  
return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle("ระบบแจ้งคะแนน")
      .addMetaTag('viewport','width=device-width , initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
} 

function getCode(code) {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var allss =ss.getSheets();  
for (var i in allss){
  var ws =ss.getSheets()[i];
  var data=ws.getDataRange().getDisplayValues().filter(row=>{
    return row[0]==code
    })
    Logger.log(data)
  if(data.length>0) break;
}

var stdCodesList = data.map (function(r) { return r[0]; }); 

var stdList = data.map(function(r) { 
  
return [`<table class="table table-bordered mt-3">  
        <thead class="p-3 mb-2 bd-blue-500 text-white">
        <tr>
        <th scope="col" colspan="12"><center><i class="fas fa-user-graduate"></i> ผลการเรียน</center></th>
        </tr>
        </thead>
        <tbody>
        <tr>
        <td colspan="12"><h3><font color="Green">${r[1]}</font></h3></td>
        </tr> 
        <tr>
        <td colspan="12"><h3><font color="Green">${r[2]}</font></h3></td>
        </tr> 
        <thead class="p-3 mb-2 bd-blue-500 text-white">
        <tr>
        <th scope="col" ><center>วิชา</center></th>
        <th scope="col"><center>เกรด</center></th>
        </tr>
        </thead>
        <tr>
        <td><b>${r[3]}</b></td>      
        <td ><center> 
        ${r[4]}</center></td>
       
        </tr> 
        <tr>
        <td><b>${r[5]}</b></td>      
        <td ><center>${r[6]}</center></td>
        
        </tr> 
        <tr>
        <td><b>${r[7]}</b></td>      
        <td ><center>${r[8]}</center></td>
        
        </tr> 
        <tr>
        <td><b>${r[9]}</b></td>      
        <td ><center>${r[10]}</center></td>
         
        </tr>                          
        <tr>
        <td><b>${r[11]}</b></td>      
        <td ><center>${r[12]}</center></td>
        
        </tr> 
        <tr>
        <td><b>${r[13]}</b></td>      
        <td ><center>${r[14]}</center></td>
       
        </tr> 
        <tr>
        <td><b>${r[15]}</b></td>      
        <td ><center>${r[16]}</center></td>
       
        </tr> 
        <tr>
        <td><b>${r[17]}</b></td>      
        <td ><center>${r[18]}</center></td>
        
        </tr> 
        <tr>
        <td><b>${r[19]}</b></td>      
        <td ><center>${r[20]}</center></td>
       
        </tr> 
        <tr>
        <td><b>${r[21]}</b></td>      
        <td ><center>${r[22]}</center></td>
       
        </tr>   
        </tr> 
        <tr>
        <td><b>${r[23]}</b></td>      
        <td ><center>${r[24]}</center></td>
       
        </tr>   
       
       <tr>
        <td><b>ดาวน์โหลด</b></td>                    
        <td colspan="10"<center><a target="_blank" href="${r[26]}" class="p-2 mb-2 bd-blue-500 text-white"><i class="fas fa-download"></i></i> เอกสาร</a></center>       
        </tr>              
        </tbody>
        </table>                   
        `];
});

var position = stdCodesList.indexOf(code); 
if(position > -1){
return stdList[position];
} else {
return '<center>*ไม่พบข้อมูล<br><img src="https://ltschool.web.app/pic/falseok.gif" width="200" height="200"></center>';
  } 
}

function getURL(){
return ScriptApp.getService().getUrl()
}
