var BATCH_SIZE = 50;
var ss = SpreadsheetApp.getActiveSpreadsheet();

function activityReport() {
    var status = ScriptProperties.getProperty("status");
    if (status == null) {
        // If the script is triggered for the first time, init
        init_();
    } else {
        status = Utilities.jsonParse(status);
        var previousMonth = new Date(new Date().getFullYear(), new Date().getMonth() - 1, 1).getMonth();
        if (status == null || (status.custom == false && status.previousMonth != previousMonth)) {
            init_();
            fetchEmails_(status.customReport, status.groupEmailAddress);
        }
        // If report not sent, continue to work on the report
        else if (status.reportSent == "no") {
            fetchEmails_(status.customReport, status.groupEmailAddress);
        }
    }
}

function fetchEmails_(customReport, groupEmailAddress) {
    var variables = Utilities.jsonParse(ScriptProperties.getProperty("variables"));
    if (!customReport) {
        var query = "after:" + variables.year + "/" + (variables.previousMonth) + "/31";
        query += " before:" + new Date().getYear() + "/" + (variables.previousMonth + 1) + "/31";
    } else {
        var query = "after:" + Utilities.formatDate(new Date(variables.startDate), variables.userTimeZone, 'yyyy/MM/dd');
        query += " before:" + Utilities.formatDate(new Date(variables.endDate), variables.userTimeZone, 'yyyy/MM/dd');
    }
    query += " to:"+groupEmailAddress;
    var startDate = new Date(variables.startDate).getTime();
    var endDate = new Date(variables.endDate).getTime();
    var conversations = GmailApp.search(query, variables.range, BATCH_SIZE);
    variables.nbrOfConversations += conversations.length;
    for (var i = 0; i < conversations.length; i++) {
        Utilities.sleep(1000);
        var conversationId = conversations[i].getId();
        var firstMessageSubject = conversations[i].getFirstMessageSubject();
        var someoneAnswered = false;
        var messages = conversations[i].getMessages();
        var nbrOfMessages = messages.length;
        variables.nbrOfMessages += nbrOfMessages;
        variables.nbrOfEmailsPerConversation[nbrOfMessages]++;
        for (var j = 0; j < 10; j++) {
            if (variables.topThreads[j][1] < nbrOfMessages) {
                variables.topThreads.splice(j, 0, [firstMessageSubject, nbrOfMessages]);
                variables.topThreads.pop();
                j = 10;
            }
        }
        var timeOfFirstMessage = 0;
        var waitingTime = 0;
        for (var j = 0; j < nbrOfMessages; j++) {
            var process = true;
            var date = messages[j].getDate();
            var month = date.getMonth();
            if (customReport) {
                if (date.getTime() < startDate || date.getTime() > endDate) {
                    process = false;
                }
            } else {
                if (month != variables.previousMonth) {
                    process = false;
                }
            }
            if (process) {
                var time = Utilities.formatDate(date, variables.userTimeZone, "H");
                var day = Utilities.formatDate(date, variables.userTimeZone, "d") - 1;
                // Use function from Utilities file
                variables = countSendsPerDaysOfWeek_(variables, date);
                var body = messages[j].getBody();
                // Use function from Utilities file
                var resultsFromCalcMessagesLength = calcMessagesLength_(variables, body);
                variables = resultsFromCalcMessagesLength[0];
                var messageLength = resultsFromCalcMessagesLength[1];
                if (j == 0) {
                    timeOfFirstMessage = date.getTime();
                } else if (j > 0) {
                    someoneAnswered = true;
                    // Use function from Utilities file
                    variables = calcWaitingTime_(variables, date, timeOfFirstMessage);
                }
                variables.timeOfEmailsReceived[time]++;
                variables.dayOfEmailsReceived[day]++;
                if (customReport) {
                    variables.monthOfEmailsReceived[month]++;
                }

            }
        }
    }
    variables.range += BATCH_SIZE;
    ScriptProperties.setProperty("variables", Utilities.jsonStringify(variables));
    if (conversations.length < BATCH_SIZE) {
        sendReport_(variables);
    }
}