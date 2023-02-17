// Calculate time elapsed from start time to current time
function isTimeUp_(start) {
  var now = new Date();
  return 'time: ' + ((now.getTime() - start.getTime()))/1000 ;
}

// Google Apps Script entry function that creates kaspi XML file and shares it
function doGet() {
  // Get the ID of the spreadsheet containing data to be used in the XML file
  var spreadsheetId =  DriveApp.getFilesByName('Итоговая Общая база').next().getId();
  
  // Open the spreadsheet and get the first sheet
  var spread1 = SpreadsheetApp.openById(spreadsheetId); 
  var sheet1 = spread1.getSheets()[0];
  
  // Declare variables to be used later
  var the_text = '';
  var line = String.fromCharCode(13)+String.fromCharCode(10);
  var start = new Date();

  
  // Add XML header and opening tag
  the_text +=  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+line;
  the_text +=  '<kaspi_catalog date="string"'+line;
  the_text +=  '              xmlns="kaspiShopping"'+line;
  the_text +=  '          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'+line;
  the_text +=  'xsi:schemaLocation="kaspiShopping http://kaspi.kz/kaspishopping.xsd">'+line;
  the_text +=  '    <company>Company_name</company>'+line;
  the_text +=  '    <merchantid>Company_name_id</merchantid>'+line;
  the_text +=  '    <offers>'+line;
  
  // Get data from specific ranges in the sheet
  var last = sheet1.getLastRow();
  var R = sheet1.getRange('R3:R'+last).getValues();    
  var E = sheet1.getRange('E3:E'+last).getValues();
  var B = sheet1.getRange('B3:B'+last).getValues();  
  var S = sheet1.getRange('S3:S'+last).getValues();    
  var J = sheet1.getRange('J3:J'+last).getValues();  
  var F = sheet1.getRange('F3:F'+last).getValues();  
  var AX = sheet1.getRange('AX3:AX'+last).getValues();  
  var AY = sheet1.getRange('AY3:AY'+last).getValues();   
  var K = sheet1.getRange('K3:K'+last).getValues();     
  
for (var i = 0; i <last-2 ; i++){  // Loop through rows in data array
  if (S[i][0] =='yes'){  // Check if the first column of the current row is equal to 'yes'
    
    // Get values from specific columns of the current row
    var sku = B[i][0];
    var model = F[i][0];
    var brand = E[i][0];
    var availability = J[i][0];
    var price = R[i][0];
    var preset_1 = AY[i][0];       
    var preset_2 = K[i][0];
    var preset_3 = AX[i][0];
        
        // offer of XML element
    var the_text +='         <offer sku="' + sku + '">' + line +
                   '            <model>' + model + '</model>' + line +
                   '            <brand>' + brand + '</brand>' + line +
                   '            <availabilities>' + line;
      
    //checks if specific preorder setting is not empty  (set)
      if (preset_1 !=''){

        
        //checks if number of stock is bigger or equal preset number for current item
        // for example if preset is 3, and stock is equal to 3 or is more, we can open preorder for this item for 4 days
        
        if (preset_2<=preset_1){
        
        the_text+='               <availability available="'+ availability +'" storeId="PP1" preOrder="4"/>'+line;
        };
        
        
      }
      else   //otherwise directly write number of days as preorder if it's not empty  
      if (preset_3 !=''){
        
      the_text+='               <availability available="'+ availability +'" storeId="PP1" preOrder="'+preset_3+'"/>'+line;
      }
      else {otherwise write default preset
               
      
    the_text+='               <availability available="'+ availability +'" storeId="PP1"/>'+line;         
    the_text+='               <availability available="no" storeId="PP2"/>'+line;
    the_text+='               <availability available="no" storeId="PP3"/>'+line;
    the_text+='               <availability available="no" storeId="PP4"/>'+line;
    the_text+='               <availability available="no" storeId="PP5"/>'+line;
        }
        
    the_text+='            </availabilities>'+line;
      
    the_text+='            <price>' +price+ '</price>'+line;
    the_text+='            </offer>'+line;
                

      
    }};    
    
  the_text+='    </offers>'+line;
  the_text+='</kaspi_catalog>'+line;    
    
// Remove any existing 'kaspi.xml' file from Drive
if (DriveApp.getFilesByName('kaspi.xml').hasNext()){
 var id= DriveApp.getFilesByName('kaspi.xml').next();  
  
   
    Drive.Files.remove(id.getId());
    
  };   
    
// Create a new 'kaspi.xml' file with the the_text
var file = DriveApp.createFile('kaspi.xml', the_text);
file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

// Log the ID of the new file and the execution time
Logger.log(isTimeUp_(start));
Logger.log(file.getId());

// Return the ID of the new file as a response
return ContentService.createTextOutput('kaspi_id: ' + file.getId());
