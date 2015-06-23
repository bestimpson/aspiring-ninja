/*
Author Bruce Stimpson (bruce.stimpson@comcast.net)
This script is container bound and is meant to
generate a Monthly Review document from a
Google Doc template using form responses submitted by employees
*/

//review variables

var baseReviewName = "Monthly Performance Review";
var reviewsFolder = DriveApp.getFolderById("[YOUR FOLDER ID HERE]");
var reviewDoc = DriveApp.getFileById("[YOUR FILE ID HERE]").makeCopy(baseReviewName, reviewsFolder);
var reviewDocId = reviewDoc.getId();
var docUrl = reviewDoc.getUrl();

//employee data

var empSheet = SpreadsheetApp.openById("[YOUR FILE ID HERE]"); //prod
SpreadsheetApp.setActiveSpreadsheet(empSheet);
var empArray = empSheet.getDataRange().getValues();

//assessment data to be populated based on user input and data lookups from user input

var timeStamp;
var employeeEmail;
var reviewPeriod;
var valuesExplanation;
var employeeStrengths;
var employeeDevelopment;
var employeeName;
var managerName;
var managerEmail;
var lvl2Email;

//tracker variables

var tracker = SpreadsheetApp.openById("[YOUR FILE ID HERE]");
var submissionArray = tracker.getSheetByName("Employees").getDataRange().getValues();
var monthColumnLookup = tracker.getSheetByName("Lookup").getDataRange().getValues();
var submissionStatusText = "Submitted";

//email notification variables

//var ccList = "email.address@domain.com"; 
var sent = "Sent";

function sendManagerNotification() {
    "use strict";
    var baseMgrSubj, mgrSubj, baseMgrMsg, mgrMsg, mgrEmailSent;
    if (tracker.getSheetByName("Notifications").getRange("A2").getValue() !== "Manager Notification") {
        throw new Error("Could not find the manager notification data. Check the Notifications tab in the 2015 Review Tracker sheet.");
    } else {
        baseMgrSubj = tracker.getSheetByName("Notifications").getRange("B2").getValue();
        mgrSubj = baseMgrSubj.replace(/<empName>/g, employeeName);
        mgrSubj = mgrSubj.replace(/<revPeriod>/g, reviewPeriod);
        baseMgrMsg = tracker.getSheetByName("Notifications").getRange("C2").getValue();
        mgrMsg = baseMgrMsg.replace(/<managerName>/g, managerName);
        mgrMsg = mgrMsg.replace(/<empName>/g, employeeName);
        mgrMsg = mgrMsg.replace(/<revPeriod>/g, reviewPeriod);
        mgrMsg = mgrMsg.replace(/<docUrl>/g, docUrl);
        mgrEmailSent = tracker.getSheetByName("Notifications").getRange("D2").getValue();
        if (mgrEmailSent !== sent) {
            MailApp.sendEmail(managerEmail, mgrSubj, mgrMsg, {
                htmlBody: mgrMsg,
                name: "Monthly Review Process",
                replyTo: employeeEmail,
                //cc: ccList
            });
            tracker.getSheetByName("Notifications").getRange(2, 4).setValue(sent);
            SpreadsheetApp.flush();
        } else {
            MailApp.sendEmail({
                to: "email.address@domain.com",
                subj: "Employee Review Submission Confirmation Delivery Failure",
                body: "Confirmation email was not sent because it may already have been sent.",
                name: "Review Process Script"
            });
        }
        tracker.getSheetByName("Notifications").getRange(2, 4).setValue("");
        SpreadsheetApp.flush();
    }
}

function sendEmployeeNotification() {
    "use strict";
    var baseEmpSubj, empSubj, baseEmpMsg, empMsg, empEmailSent;
    if (tracker.getSheetByName("Notifications").getRange("A1").getValue() !== "Employee Notification") {
        throw new Error("Could not find the employee notification data. Check the Notifications tab in the 2015 Review Tracker sheet.");
    } else {
        baseEmpSubj = tracker.getSheetByName("Notifications").getRange("B1").getValue();
        empSubj = baseEmpSubj.replace(/<revPeriod>/g, reviewPeriod);
        baseEmpMsg = tracker.getSheetByName("Notifications").getRange("C1").getValue();
        empMsg = baseEmpMsg.replace(/<empName>/g, employeeName);
        empMsg = empMsg.replace(/<revPeriod>/g, reviewPeriod);
        empMsg = empMsg.replace(/<docUrl>/g, docUrl);
        empEmailSent = tracker.getSheetByName("Notifications").getRange("D1").getValue();
        if (empEmailSent !== sent) {
            MailApp.sendEmail(employeeEmail, empSubj, empMsg, {
                htmlBody: empMsg,
                name: "Monthly Review Process",
                replyTo: managerEmail
            });
            tracker.getSheetByName("Notifications").getRange(1, 4).setValue(sent);
            SpreadsheetApp.flush();
        } else {
            MailApp.sendEmail({
                to: "email.address@domain.com",
                subj: "Employee Review Submission Confirmation Delivery Failure",
                body: "Confirmation email was not sent because it may already have been sent.",
                name: "Review Process Script"
            });
        }
        tracker.getSheetByName("Notifications").getRange(1, 4).setValue("");
        SpreadsheetApp.flush();
    }
    
    sendManagerNotification();
    
}

function updateReviewTracker() {
    "use strict";
    var i, column, row = 1, trackerUpdateFailSubj, trackerUpdateFailMsg;
    for (i = 0; i < monthColumnLookup.length; i += 1) {
        if (monthColumnLookup[i][0] === reviewPeriod) {
            column = monthColumnLookup[i][1];
            break;
        } else if (i === (monthColumnLookup.length - 1)) {
            MailApp.sendEmail("email.address@domain.com", "Monthly Self-Assessment Tracker Not Updated", "The submission was not for a month in 2015");
        }
    }
    for (i = 0; i < submissionArray.length; i += 1) {
        if (submissionArray[i][0] === employeeEmail) {
            row += i;
            tracker.setActiveSheet(tracker.getSheets()[0]).getRange(row, column).setValue(submissionStatusText);
            break;
        } else if (i === (submissionArray.length - 1)) {
            trackerUpdateFailSubj = "Review Submission Tracker Update Failure";
            trackerUpdateFailMsg = employeeEmail + " submitted a self-assessment for " + reviewPeriod + " but the submission tracker update failed.";
            MailApp.sendEmail("email.address@domain.com", trackerUpdateFailSubj, trackerUpdateFailMsg);
        }
    }
    
    sendEmployeeNotification();
    
}

function saveReview() {
    "use strict";
    var finishedReview = DocumentApp.openById(reviewDocId);
    finishedReview.saveAndClose();
    
    updateReviewTracker();
}

function setReviewPermissions() {
    "use strict";
    var review = DocumentApp.openById(reviewDocId), viewers = review.getViewers(), editors = review.getEditors(), i, j, emailAddress;
    for (i = 0; i < viewers.length; i += 1) {
        emailAddress = viewers[i].getEmail();
        if (emailAddress !== "") {
            review.removeViewer(emailAddress);
        }
    }
    for (j = 0; j < editors.length; j += 1) {
        emailAddress = editors[j].getEmail();
        if (emailAddress !== "") {
            review.removeEditor(emailAddress);
        }
    }
    review.addViewer(employeeEmail);
    review.addEditor(managerEmail);
    review.addEditor(lvl2Email);
    
    saveReview();
}

function addCustomProperties(fileId) {
    "use strict";
    var employeeInfo, managerInfo, reviewPeriodInfo, emailSubject, emailMessage, reviewNoticeSubject = tracker.getSheetByName("Notifications").getRange("B3").getValue(), reviewNoticeMessage = tracker.getSheetByName("Notifications").getRange("C3").getValue();
    if (tracker.getSheetByName("Notifications").getRange("A3").getValue() !== "Feedback Notification") {
        throw new Error("Could not find the feedback notification content. Check cell A3 on the Notifications tab in the 2015 Review Tracker sheet.");
    } else {
        employeeInfo = {
            key: "employeeEmail",
            value: employeeEmail,
            visibility: "PUBLIC"
        };
        managerInfo = {
            key: "managerName",
            value: managerName,
            visibility: "PUBLIC"
        };
        reviewPeriodInfo = {
            key: "reviewPeriod",
            value: reviewPeriod,
            visibility: "PUBLIC"
        };
        emailSubject = {
            key: "Subject",
            value: reviewNoticeSubject,
            visibility: "PUBLIC"
        };
        emailMessage = {
            key: "Message",
            value: reviewNoticeMessage,
            visibility: "PUBLIC"
        };
        Drive.Properties.insert(employeeInfo, reviewDocId);
        Drive.Properties.insert(managerInfo, reviewDocId);
        Drive.Properties.insert(reviewPeriodInfo, reviewDocId);
        Drive.Properties.insert(emailSubject, reviewDocId);
        Drive.Properties.insert(emailMessage, reviewDocId);
    }
    
    setReviewPermissions();
}

function renameReviewDoc() {
    "use strict";
    reviewDoc.setName(baseReviewName + " for " + employeeName + " - " + reviewPeriod);
    
    addCustomProperties();
}

function replaceDocumentText() {
    "use strict";
    var body = DocumentApp.openById(reviewDocId).getBody();
    body.replaceText("keyEmpName", employeeName);
    body.replaceText("keyRevPeriod", reviewPeriod);
    body.replaceText("keyManagerName", managerName);
    body.replaceText("keyValExplanation", valuesExplanation);
    body.replaceText("keyEmpStrengths", employeeStrengths);
    body.replaceText("keyEmpDevelopment", employeeDevelopment);
    
    renameReviewDoc();
    
}

function createAssessmentData(e) {
    
    //assign variable values using data from form submision
    
    "use strict";
    timeStamp = e.values[0];
    employeeEmail = e.values[1];
    reviewPeriod = e.values[2];
    valuesExplanation = e.values[3]
    employeeStrengths = e.values[4];
    employeeDevelopment = e.values[5];
    
    //Loop through data in Employee roster to assign remaining values
    
    var i;
    for (i = 0; i < empArray.length; i += 1) {
        if (empArray[i][0] === employeeEmail) {
            employeeName = empArray[i][1];
            managerName = empArray[i][2];
            managerEmail = empArray[i][3];
            lvl2Email = empArray[i][5];
            break;
        } else if (i === (empArray.length - 1)) {
            employeeName = employeeEmail;
            managerName = "Default Manager"; //specify an individual who can route the review to the appropriate manager
            managerEmail = "email.address@domain.com";
            lvl2Email = "email.address@domain.com";
        }
    }
    
    replaceDocumentText();
    
}
