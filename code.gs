var Lookuptable_InvestmentGoal = {
  "create a financial emergency fund": 0,
  "save for future retirement": 1,
  "fund my children's education": 2,
  "save for a down-payment on a home": 3,
  "generate income": 4,
  "support for charity": 5,
  "support my heirs": 6,
};

var Lookuptable_RiskTaker = {
  "plays it safe, avoids risk": 0,
  "willing to take risks after completing adequate research": 1,
  "a real gambler": 2,
};

var Lookuptable_TvGameShow = {
  "A 5% chance of winning $100,000": 0,
  "A 50% chance of winning $5,000": 1,
};

var sheet = SpreadsheetApp.getActiveSheet();
var textValues = sheet.getRange("A:AB").getValues();

// for (var i = 0; i < 20; i++) {
//     var textValue = textValues[i][5];
//     console.log(textValue)
//     var numericValue = Lookuptable_RiskTaker[textValue];
//     console.log(numericValue)
//   }

function ConvertNumeric(id, column) {
  var textValue = textValues[id][column];

  console.log(textValue);
  if (column === 4) {
    var arr = textValue.split(", ");
    var numericValue = [];
    for (j = 0; j <= arr.length; j++) {
      console.log(arr[j]);
      numericValue.push(Lookuptable_InvestmentGoal[arr[j]]);
    }
  } else if (column === 5) {
    var numericValue = Lookuptable_RiskTaker[textValue];
  } else if (column === 6) {
    var numericValue = Lookuptable_TvGameShow[textValue];
  }
  console.log(numericValue);
}

ConvertNumeric(2, 4);

function makedoc(id) {
  // var doc = DocumentApp.openById("1i4WSnzaTpMchUcjzF85rbRRpWbJ4lFOl_PfDfY13C58");

  // body.replaceText('%InvestorName%', investorName);
  var docx = DriveApp.getFileById(
    "1i4WSnzaTpMchUcjzF85rbRRpWbJ4lFOl_PfDfY13C58"
  );

  var copy = docx.makeCopy("Investment Policy");
  var doc = DocumentApp.openById(copy.getId());

  var body = doc.getBody();
  // var text = body.getText();

  body.replaceText("%InvestorName%", textValues[id][1]);
  body.replaceText("%Date%", textValues[id][0]);
  body.replaceText("%Description%", textValues[id][27]);
  body.replaceText("%Goal1%", textValues[id][4].split(", ")[0]);
  body.replaceText("%Goal2", textValues[id][4].split(", ")[1]);
  body.replaceText("%Goal3%", textValues[id][4].split(", ")[2]);
  body.replaceText("%RiskTolerance%", textValues[id][5]);

  doc.saveAndClose();

  var emailID = "akshat15599@gmail.com";
  var subject = "Regarding your investment";
  var emailBody = "This is demo email from markdale.";
  var attach = DriveApp.getFileById(copy.getId());
  var pdfattach = attach.getAs(MimeType.PDF);
  MailApp.sendEmail(emailID, subject, emailBody, { attachments: [pdfattach] });
}

for (i = 0; i < 20; i++) {
  ConvertNumeric(i, 4), ConvertNumeric(i, 5), ConvertNumeric(i, 6), makedoc(i);
}
