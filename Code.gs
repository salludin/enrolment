function myFunction(e) {
var sheet = SpreadsheetApp.getActiveSheet();
var row =  SpreadsheetApp.getActiveSheet().getLastRow();
var level = e.values[1];
var grade = e.values [2];
var country = e.values[9];
if (grade == "Grade 1"){
  var academicyear = e.values[4];
}else{
  var academicyear = e.values[3];
}
var emailadmission = "admission@mutiaraharapan.sch.id";
var apiKey = 'xnd_production_e3nYyzyZtJCrnoowK6BsvwL0dzbklIHO9r20qEfXafB12IFj9JQHcEtaVqES';
  var Basic = Utilities.base64Encode(apiKey + ':0');
  var header = {
   "Authorization" : "Basic " + Basic
  };
if (level == 'Playgroup' || level == 'Kindergarten') {
var no = 300000 + 3650;
sheet.getRange(row,23).setNumberFormat("#,##0").setValue(no);
 var biaya = sheet.getRange(row, 23).getDisplayValue();
 var studentname = e.values[13];
 var email = e.values[7];
 var desc = "Form PSB " + level + " - " + studentname + "";
 var formDataInvoice = {
  "external_id": "PSB " + academicyear + " - " + row + "",
	"amount": no,
	"payer_email": email,
	"description": desc,
  "invoice_duration": 604800,
  "reminder_time": 3
    
  }
  var optionsInvoice = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataInvoice)
  };
  var CreateInvoice = UrlFetchApp.fetch("https://api.xendit.co/v2/invoices", optionsInvoice);
  var json = CreateInvoice.getContentText();
  var data = JSON.parse(json);
  var link = data["invoice_url"];
  var id = "PSB " + academicyear + " - " + row + "";
  var id_invoice = sheet.getRange(row,22).setValue(id);
  var BodyTemplate = HtmlService.createTemplateFromFile("email");
  BodyTemplate.studentname = studentname;
  BodyTemplate.fee = "300,000";
  BodyTemplate.adm = "3,650";
  BodyTemplate.total = biaya;
  BodyTemplate.link = link;
  BodyTemplate.level = level;
  var HtmlforBody = BodyTemplate.evaluate().getContent();
}else if (level == 'Primary') {
  var no = 500000 + 3650;
sheet.getRange(row,23).setNumberFormat("#,##0").setValue(no);
 var biaya = sheet.getRange(row, 23).getDisplayValue();
 var studentname = e.values[13];
 var email = e.values[7]; 
 var desc = "Form PSB Primary - " + studentname + "";
 var formDataInvoice = {
  "external_id": "PSB " + academicyear + " - " + row + "",
	"amount": no,
	"payer_email": email,
	"description": desc,
  "invoice_duration": 604800,
  "reminder_time": 3
    
  }
  var optionsInvoice = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataInvoice)
  };
  var CreateInvoice = UrlFetchApp.fetch("https://api.xendit.co/v2/invoices", optionsInvoice);
  var json = CreateInvoice.getContentText();
  var data = JSON.parse(json);
  var link = data["invoice_url"];
  var id = "PSB " + academicyear + " - " + row + "";
  var id_invoice = sheet.getRange(row,22).setValue(id);
  var BodyTemplate = HtmlService.createTemplateFromFile("email");
  BodyTemplate.studentname = studentname;
  BodyTemplate.fee = "500,000";
  BodyTemplate.adm = "3,650";
  BodyTemplate.total = biaya;
  BodyTemplate.link = link;
  BodyTemplate.level = "Primary";
  var HtmlforBody = BodyTemplate.evaluate().getContent();
}else if (level == 'Lower Secondary') {
  var no = 450000 + 3650;
sheet.getRange(row,23).setNumberFormat("#,##0").setValue(no);
 var biaya = sheet.getRange(row, 23).getDisplayValue();
 var studentname = e.values[13];
 var email = e.values[7]; 
 var desc = "Form PSB Lower Secondary - " + studentname + "";
 var formDataInvoice = {
  "external_id": "PSB " + academicyear + " - " + row + "",
	"amount": no,
	"payer_email": email,
	"description": desc,
  "invoice_duration": 604800,
  "reminder_time": 3
    
  }
  var optionsInvoice = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataInvoice)
  };
  var CreateInvoice = UrlFetchApp.fetch("https://api.xendit.co/v2/invoices", optionsInvoice);
  var json = CreateInvoice.getContentText();
  var data = JSON.parse(json);
  var link = data["invoice_url"];
  var id = "PSB " + academicyear + " - " + row + "";
  var id_invoice = sheet.getRange(row,22).setValue(id);
  var BodyTemplate = HtmlService.createTemplateFromFile("email");
  BodyTemplate.studentname = studentname;
  BodyTemplate.fee = "450,000";
  BodyTemplate.adm = "3,650";
  BodyTemplate.total = biaya;
  BodyTemplate.link = link;
  BodyTemplate.level = "Lower Secondary";
  var HtmlforBody = BodyTemplate.evaluate().getContent();
} else if (level == 'Upper Secondary') {
  var no = 450000 + 3650;
sheet.getRange(row,23).setNumberFormat("#,##0").setValue(no);
 var biaya = sheet.getRange(row, 23).getDisplayValue();
 var studentname = e.values[13];
 var email = e.values[7]; 
 var desc = "Form PSB Upper Secondary - " + studentname + "";
 var formDataInvoice = {
  "external_id": "PSB " + academicyear + " - " + row + "",
	"amount": no,
	"payer_email": email,
	"description": desc,
  "invoice_duration": 604800,
  "reminder_time": 3
    
  }
  var optionsInvoice = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataInvoice)
  };
  var CreateInvoice = UrlFetchApp.fetch("https://api.xendit.co/v2/invoices", optionsInvoice);
  var json = CreateInvoice.getContentText();
  var data = JSON.parse(json);
  var link = data["invoice_url"];
  var id = "PSB " + academicyear + " - " + row + "";
  var id_invoice = sheet.getRange(row,22).setValue(id);
  var BodyTemplate = HtmlService.createTemplateFromFile("email");
  BodyTemplate.studentname = studentname;
  BodyTemplate.fee = "450,000";
  BodyTemplate.adm = "3,650";
  BodyTemplate.total = biaya;
  BodyTemplate.link = link;
  BodyTemplate.level = "Upper Secondary";
  var HtmlforBody = BodyTemplate.evaluate().getContent();
} else if (level == 'Primary Development Class (Special Education Program)') {
  var no = 1250000 + 3650;
sheet.getRange(row,23).setNumberFormat("#,##0").setValue(no);
 var biaya = sheet.getRange(row, 23).getDisplayValue();
 var studentname = e.values[13];
 var email = e.values[7]; 
 var desc = "Form PSB Primary Development Class - " + studentname + "";
 var formDataInvoice = {
  "external_id": "PSB " + academicyear + " - " + row + "",
	"amount": no,
	"payer_email": email,
	"description": desc,
  "invoice_duration": 604800,
  "reminder_time": 3
    
  }
  var optionsInvoice = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataInvoice)
  };
  var CreateInvoice = UrlFetchApp.fetch("https://api.xendit.co/v2/invoices", optionsInvoice);
  var json = CreateInvoice.getContentText();
  var data = JSON.parse(json);
  var link = data["invoice_url"];
  var id = "PSB " + academicyear + " - " + row + "";
  var id_invoice = sheet.getRange(row,22).setValue(id);
  var BodyTemplate = HtmlService.createTemplateFromFile("email");
  BodyTemplate.studentname = studentname;
  BodyTemplate.fee = "1,250,000";
  BodyTemplate.adm = "3,650";
  BodyTemplate.total = biaya;
  BodyTemplate.link = link;
  BodyTemplate.level = "Primary Development Class";
  var HtmlforBody = BodyTemplate.evaluate().getContent();
} else if (level == 'Lower Secondary Development Class (Special Education Program)') {
  var no = 1250000 + 3650;
sheet.getRange(row,23).setNumberFormat("#,##0").setValue(no);
 var biaya = sheet.getRange(row, 23).getDisplayValue();
 var studentname = e.values[13];
 var email = e.values[7]; 
 var desc = "Form PSB Lower Secondary Development Class - " + studentname + "";
 var formDataInvoice = {
  "external_id": "PSB " + academicyear + " - " + row + "",
	"amount": no,
	"payer_email": email,
	"description": desc,
  "invoice_duration": 604800,
  "reminder_time": 3
    
  }
  var optionsInvoice = {
  'method' : 'post',
  'headers' :header,
  'contentType': 'application/json',
  // Convert the JavaScript object to a JSON string.
  'payload' : JSON.stringify(formDataInvoice)
  };
  var CreateInvoice = UrlFetchApp.fetch("https://api.xendit.co/v2/invoices", optionsInvoice);
  var json = CreateInvoice.getContentText();
  var data = JSON.parse(json);
  var link = data["invoice_url"];
  var id = "PSB " + academicyear + " - " + row + "";
  var id_invoice = sheet.getRange(row,22).setValue(id);
  var BodyTemplate = HtmlService.createTemplateFromFile("email");
  BodyTemplate.studentname = studentname;
  BodyTemplate.fee = "1,250,000";
  BodyTemplate.adm = "3,650";
  BodyTemplate.total = biaya;
  BodyTemplate.link = link;
  BodyTemplate.level = "Lower Secondary Development Class";
  var HtmlforBody = BodyTemplate.evaluate().getContent();
}
var subject = "Enrolment Payment of " + studentname + " - Mutiara Harapan Islamic School";
var options = {};
  options.htmlBody = HtmlforBody;
GmailApp.sendEmail("" +email+ "", 
                   subject,
                   "Please open this email with support HTML",
                   options
                  );
if (country == "Indonesia"){
  zipcode = e.values[10];
}else{
  zipcode = e.values[12];
}
var Geocodingcurrent = UrlFetchApp.fetch("https://maps.googleapis.com/maps/api/geocode/json?address=" + zipcode +"+Indonesia&key=AIzaSyDW0XaRzzDDMgUdPOXvPYrKej__-b6Cby4");
var jsoncurrent = Geocodingcurrent.getContentText();
var datacurrent = JSON.parse(jsoncurrent);
  var filtered_kelurahan = datacurrent.results[0].address_components.filter(function(address_component){
    return address_component.types.includes("administrative_area_level_4");
      })
  var filtered_kecamatan = datacurrent.results[0].address_components.filter(function(address_component){
    return address_component.types.includes("administrative_area_level_3");
      })
  var filtered_kota = datacurrent.results[0].address_components.filter(function(address_component){
    return address_component.types.includes("administrative_area_level_2");
      })
  var kota = filtered_kota.length ? filtered_kota[0].long_name: "";
  var kecamatan = filtered_kecamatan.length ? filtered_kecamatan[0].long_name: "";
  var kelurahan = filtered_kelurahan.length ? filtered_kelurahan[0].long_name: "";
  var setkelurahan = sheet.getRange(row,26).setValue(kelurahan);
  var setkecamatan = sheet.getRange(row,25).setValue(kecamatan);
  var setkota = sheet.getRange(row,24).setValue(kota);



}