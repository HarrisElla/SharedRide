var WHATSAPP_MESSAGE = "היי, מחכים לך בקבוצת הוואטסאפ של ההסעה: ";
var WHATSAPP_MESSAGE2 = "היי, ראיתי שנרשמת להסעה";
var GOOGLE_FORMS_SHEET_NAME = "תגובות לטופס 1";
var PAYBOX_SHEET_NAME = "paybox";
var PHONE_NUMBER_COLUMN_NAME = "טלפון סלולרי:";
var PHONE_IN_PAYBOX_COLUMN_NAME = "פלאפון"
var MERGED = "merged";
var WHATSAPP_COLUMN_NAME = "whatsapp";
var WHATSAPP_COLUMN_NAME2 = "whatsapp2";
var SHEET = SpreadsheetApp.getActiveSpreadsheet();
var Formsheet;
var Payboxsheet;
var WhatsappGroups = [];
var GOOGLE_FORM_CODE;

function showError(msg) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Error', msg, ui.ButtonSet.OK);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
    .addItem("פיצול טאבים לפי יום",'splitToTabs')
    .addItem("הוספת לינקים לקבוצת וואטסאפ",'addWhatsAppLinks')
    .addItem('השוואה לפייבוקס', 'mergeSheets')
    .addToUi();
}

function getDaysFromForm() {
  // Replace 'FORM_ID' with the actual Form ID.
  var form = FormApp.openById(GOOGLE_FORM_CODE);

  // Replace 'QUESTION_INDEX' with the index of the question you want to retrieve options from.
  var questionIndex = 0;

  // Get the question object
  var question = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE)[questionIndex];

  // Get the options from the question
  var options = question.asMultipleChoiceItem().getChoices();

  // Print the options to the current sheet
  var days = [];
  for (var i = 0; i < options.length; i++) {
    days[i] = options[i].getValue();
  }
  console.log(days);
  return days;
}

function splitToTabs() {
  if (!fetchData()) return;
  var days = getDaysFromForm();
  var formValues = getGoogleFormContent();
  var headers = Formsheet.getRange(1, 1, 1, Formsheet.getLastColumn())
  var tabs = [];
  var lastRows = [];
  var lastTimeStamp = []
  var timestampIndex = formValues[0].indexOf("חותמת זמן");
  var groupIndex = formValues[0].indexOf(TABS_FIELD);
  

  for(var i = 0; i < days.length; i++) {
    tabs[i] = SHEET.getSheetByName(days[i]);
    if(!tabs[i]){
      tabs[i] = SHEET.insertSheet(days[i]);
      // The code below copies the hedears
      headers.copyTo(tabs[i].getRange(1, 1));
    }
    lastRows[i] = tabs[i].getLastRow();
    lastTimeStamp[i] = tabs[i].getRange(tabs[i].getLastRow(), 1, 1, tabs[i].getLastColumn()).getValues()[0][timestampIndex]
  }
  
  console.log("ts array: " + lastTimeStamp);
  max_ts = Math.max(...lastTimeStamp);
  console.log("ts: " + max_ts);

  for (var i = 1; i < formValues.length; i++) {
    //console.log(i);
    var row = Formsheet.getRange(i+1, 1, 1, Formsheet.getLastColumn())
    row_ts = formValues[i][timestampIndex];
    if (max_ts && row_ts <= max_ts) {
      continue;
    }
    //console.log(groupIndex);

    for(var j = 0; j < days.length; j++) {
      //console.log(formValues[i][groupIndex])
      if(formValues[i][groupIndex].toString().includes(days[j])){
        //console.log(days[j]);
        console.log("add row - " + tabs[j].getLastRow() + " to tab " + days[j]);
        row.copyTo(tabs[j].getRange(tabs[j].getLastRow()+1, 1));
      }
    }
  }

  for(var j = 0; j < days.length; j++) {
    addWhatsAppStartConversationLinks(tabs[j],lastRows[j]);
  }
}


function addWhatsAppStartConversationLinks(sheet,oldLastRow = 1) {
  
  if (!fetchData()) return;
  if (!sheet) {
    sheet = SHEET.getSheetByName(GOOGLE_FORMS_SHEET_NAME);
  }
  
  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Get all the values from form sheets based on the last columns
  var lastColumn = sheet.getLastColumn();
  
  var phoneIndexForm = headers.indexOf(PHONE_NUMBER_COLUMN_NAME) + 1;

  // Create a WhatsApp column if it does not exist in the first sheet
  var whatsappIndex2 = headers.indexOf(WHATSAPP_COLUMN_NAME2) + 1;
  if (whatsappIndex2 === 0) {
    sheet.getRange(1, lastColumn + 1).setValue(WHATSAPP_COLUMN_NAME2);
    whatsappIndex2 = lastColumn + 1;
    lastColumn++;
  }
  
  lastRow = sheet.getLastRow();
  for (var j = oldLastRow+1; j <= lastRow; j++) {
    var phone = sheet.getRange(j, phoneIndexForm).getValue();
    phone = normalizePhoneNumber(phone);
    // Add the WhatsApp invitation link
   
    var whatsappUrl = `https://api.whatsapp.com/send?phone=972${phone}&text=${encodeURIComponent(WHATSAPP_MESSAGE2)}`;
    sheet.getRange(j, whatsappIndex2).setValue(whatsappUrl);
  }
}
function addGroupInviteLink(sheet, group = WHATSAPP_GROUP_1,oldLastRow = 1) {
  if (!fetchData()) return;
  if (!sheet) {
    sheet = SHEET.getSheetByName(GOOGLE_FORMS_SHEET_NAME);
  }
  
  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Get all the values from form sheets based on the last columns
  var lastColumn = sheet.getLastColumn();
  
  var phoneIndexForm = headers.indexOf(PHONE_NUMBER_COLUMN_NAME) + 1;

  // Create a WhatsApp column if it does not exist in the first sheet
  var whatsappIndex1 = headers.indexOf(WHATSAPP_COLUMN_NAME) + 1;
  if (whatsappIndex1 === 0) {
    sheet.getRange(1, lastColumn + 1).setValue(WHATSAPP_COLUMN_NAME);
    whatsappIndex1 = lastColumn + 1;
    lastColumn++;
  }
  
  var whatsappMessage1 = WHATSAPP_MESSAGE + group;
  lastRow = sheet.getLastRow();
  for (var j = oldLastRow+1; j <= lastRow; j++) {
    var phone = sheet.getRange(j, phoneIndexForm).getValue();
    phone = normalizePhoneNumber(phone);
    // Add the WhatsApp invitation link
   
    var whatsappUrl = `https://api.whatsapp.com/send?phone=972${phone}&text=${encodeURIComponent(whatsappMessage1)}`;
    sheet.getRange(j, whatsappIndex1).setValue(whatsappUrl);
  }
}

function addWhatsAppLinks() {
   if (!fetchData()) return;
  var whatsappSheet = SHEET.getSheetByName(WHATSAPP_SHEET_NAME);
  var tabs = whatsappSheet.getRange(2,1,whatsappSheet.getLastRow()-1).getValues();
  var whatsappGroups = whatsappSheet.getRange(2,2,whatsappSheet.getLastRow()-1).getValues();

  for(var j = 0; j < tabs.length; j++) {
    var sheet = SHEET.getSheetByName(tabs[j])
    addGroupInviteLink(sheet, whatsappGroups[j]);
  }

  return;
}

function mergeSheets() {
  if (!fetchData()) return;
  
  //Get all paybox sheet varaiables
  var payboxValues = getPayboxContent();
  var mergedIndexPaybox = payboxValues[0].indexOf(MERGED);
  var lastColumnPaybox = mergedIndexPaybox === -1 ? Payboxsheet.getLastColumn() : mergedIndexPaybox;
  var phoneIndexPaybox = payboxValues[0].indexOf(PHONE_IN_PAYBOX_COLUMN_NAME);
   if (mergedIndexPaybox === -1) {
    // Create a merged column if it does not exist
    Payboxsheet.getRange(1, lastColumnPaybox + 1).setValue(MERGED);
    mergedIndexPaybox = lastColumnPaybox;
    lastColumnPaybox++;
  }

   
  var tabs = getDaysFromForm();
  //var tabs = [GOOGLE_FORMS_SHEET_NAME]
  for(l=0; l< tabs.length; l++) {
    console.log("tab: " + l+1 + " of " + tabs.length);
    //Get all registration sheets varaiables
    var tab = SHEET.getSheetByName(tabs[l]);
    var formValues = tab.getRange(1, 1, tab.getLastRow(), tab.getLastColumn()).getValues();
    var mergedIndexFrom = formValues[0].indexOf(MERGED);
    var lastColumnForm = mergedIndexFrom === -1 ? tab.getLastColumn() : mergedIndexFrom;
    var phoneIndexForm = formValues[0].indexOf(PHONE_NUMBER_COLUMN_NAME);
    if (mergedIndexFrom === -1) {
      // Create a merged column if it does not exist
      tab.getRange(1, lastColumnForm + 1).setValue(MERGED);
      mergedIndexFrom = lastColumnForm;
      lastColumnForm++;
    }
    // Keep track of merged and duplicate phone numbers
    var mergedPhones = {};
    for (var i = 1; i < payboxValues.length; i++) {
      var phone2 = payboxValues[i][phoneIndexPaybox];
      phone2 = normalizePhoneNumber(phone2);
      // Loop through each row in the first sheet
      for (var j = 1; j < formValues.length; j++) {
        var phone1 = formValues[j][phoneIndexForm];
        phone1 = normalizePhoneNumber(phone1);
        var merged1 = formValues[j][mergedIndexFrom];


        // Check if phone numbers match and the row is not already merged
        if (phone1 && phone1 === phone2) {
          // Merge the rows: Append columns from the second sheet to the first
          console.log("BEEN HERE ", phone1, " mergedIndexFrom " + mergedIndexFrom);
          // Check if this phone number has already been merged
          if (mergedPhones[phone2]) {
            // Highlight the row in yellow and write 'duplicate' in the 'merged' column
            Payboxsheet.getRange(i + 1, 1, 1, lastColumnPaybox).setBackground('orange');
            Payboxsheet.getRange(i + 1, lastColumnPaybox + 1).setValue('DUPLICATE');
            break;
          } else {
            if (merged1 !== "TRUE") {
              for (var k = 0; k < payboxValues[0].length; k++) {
                tab.getRange(j + 1, lastColumnForm + k + 1).setValue(payboxValues[i][k]);
              }

              // Update the "merged" column in both sheets
              tab.getRange(j + 1, mergedIndexFrom + 1).setValue("TRUE");
              Payboxsheet.getRange(i + 1, mergedIndexPaybox + 1).setValue("TRUE");


              // Highlight the merged row in the first sheet
              tab.getRange(j + 1, 1, 1, formValues[0].length).setBackground("yellow");

            }
          }
          // Mark this phone number as merged
          mergedPhones[phone2] = true;

          break; // Break the inner loop as we've found a match
        }
      }
    }
  }
 
  

  
}

function normalizePhoneNumber(phone) {
  // Remove all non-digit characters
  var normalized = phone.replace(/\D/g, '');

  // Remove country and area codes (assuming they take up the first 2 to 8 digits)
  // Adjust this as per your specific requirements
  normalized = normalized.substring(normalized.length - 9); // Keeping last 7 digits

  return normalized;
}

function findCorrespondingValue(sheet, searchValue) {

  // Get all values from column A and column B
  var columnA = sheet.getRange("A:A").getValues();
  var columnB = sheet.getRange("B:B").getValues();

  // Find the corresponding value in column B
  for (var i = 0; i < columnA.length; i++) {
    if (columnA[i][0] === searchValue) {
      return columnB[i][0];
    }
  }

  // Return null if the searchValue is not found in column A
  return null;
}

function getPayboxContent() {
  Payboxsheet = SHEET.getSheetByName(PAYBOX_SHEET_NAME);
  if (!Payboxsheet) {
    Payboxsheet = SHEET.insertSheet(PAYBOX_SHEET_NAME);
    showError("חסר טאב פייבוקס. מסיים");
    return;
  }
  var lastColumn = Payboxsheet.getLastColumn();
  if (lastColumn < 1) {
    showError("הטאב של פייבוקס ריק. מסיים");
    return;
  }
  var values = Payboxsheet.getRange(1, 1, Payboxsheet.getLastRow(), lastColumn).getValues();
  return values;
}

function getGoogleFormContent() {
  Formsheet = SHEET.getSheetByName(GOOGLE_FORMS_SHEET_NAME);
  var lastColumn1 = Formsheet.getLastColumn();
  var values1 = Formsheet.getRange(1, 1, Formsheet.getLastRow(), Formsheet.getLastColumn()).getValues();
  return values1;
}



function fetchData() {
  var varSheet = SHEET.getSheetByName("הנחיות והגדרות");
  if (!varSheet) {
    showError("בבקשה להעתיק את טאב ההגדרות וההנחיות מקובץ אחר.");
    return false;
  }
  
  
  GOOGLE_FORMS_SHEET_NAME = findCorrespondingValue(varSheet, "שם טבלת הפורמס");
  PAYBOX_SHEET_NAME = findCorrespondingValue(varSheet, "שם טבלת הפייבוקס");
  WHATSAPP_SHEET_NAME = findCorrespondingValue(varSheet, "טאב להגדרת קבוצות");
  PHONE_NUMBER_COLUMN_NAME = findCorrespondingValue(varSheet, "שם עמודת טלפון(מדויק)");
  MERGED = findCorrespondingValue(varSheet, "שם עמודת סטטוס מיזוג");

  WHATSAPP_COLUMN_NAME = findCorrespondingValue(varSheet, "שם עמודת וואטסאפ ראשונה");
  WHATSAPP_COLUMN_NAME2 = findCorrespondingValue(varSheet, "שם עמודת וואטסאפ שניה");
  WHATSAPP_MESSAGE = findCorrespondingValue(varSheet, "הודעת וואטסאפ 1");
  WHATSAPP_MESSAGE2 = findCorrespondingValue(varSheet, "הודעת וואטסאפ 2");
  TABS_FIELD = findCorrespondingValue(varSheet, "שם עמודת פיצול לטאבים");
  GOOGLE_FORM_CODE = findCorrespondingValue(varSheet, "מפתח לטופס");
  
  
  return true;

}
