function init_(groupEmailAddress, sDate, eDate) {
    var dayOfEmailsReceived = [];
    var timeOfEmailsReceived = [];
    for (i = 0; i < 24; i++) {
        timeOfEmailsReceived[i] = 0;
    }
    var dayOfWeek = [];
    for (i = 0; i < 7; i++) {
        dayOfWeek[i] = 0;
    }
    var nbrOfEmailsPerConversation = [];
    for (i = 0; i < 101; i++) {
        nbrOfEmailsPerConversation[i] = 0;
    }
    var timeBeforeFirstResponse = [];
    for (i = 0; i < 6; i++) {
        timeBeforeFirstResponse[i] = 0;
    }
    var messagesLength = [];
    for (i = 0; i < 6; i++) {
        messagesLength[i] = 0;
    }
    var topThreads = [];
    for (i = 0; i < 10; i++) {
        topThreads.push(['', 0]);
    }

    var userTimeZone = CalendarApp.getDefaultCalendar().getTimeZone();
    var user = Session.getEffectiveUser().getEmail();
    var variables = {
        range: 0,
        nbrOfConversations: 0,
        nbrOfMessages: 0,
        nbrOfEmailsPerConversation: nbrOfEmailsPerConversation,
        timeOfEmailsReceived: timeOfEmailsReceived,
        dayOfWeek: dayOfWeek,
        timeBeforeFirstResponse: timeBeforeFirstResponse,
        messagesLength: messagesLength,
        topThreads: topThreads,
        userTimeZone: userTimeZone,
        user: user
    };
    if (sDate != undefined) {
        variables['type'] = 'custom';
        variables['startDate'] = sDate;
        variables['endDate'] = eDate;
        for (i = 0; i < 31; i++) {
            dayOfEmailsReceived[i] = 0;
        }
        var monthOfEmailsReceived = [];
        for (i = 0; i < 12; i++) {
            monthOfEmailsReceived[i] = 0;
        }
        variables['monthOfEmailsReceived'] = monthOfEmailsReceived;
        var status = {
            groupEmailAddress: groupEmailAddress,
            customReport: true,
            reportSent: "no",
        };
    }
    else {
        // Find previous month...
        variables['previous'] = new Date(new Date().getFullYear(), new Date().getMonth() - 1, 2);
        variables['previousMonth'] = variables['previous'].getMonth();
        variables['year'] = variables['previous'].getYear();
        variables['lastDay'] = daysInMonth_(variables['previousMonth'], variables['year']);
        for (i = 0; i < variables['lastDay']; i++) {
            dayOfEmailsReceived[i] = 0;
        }
        var status = {
            groupEmailAddress: groupEmailAddress,
            customReport: false,
            previousMonth: variables['previousMonth'],
            reportSent: "no"
        };
    }
    variables['dayOfEmailsReceived'] = dayOfEmailsReceived;
    ScriptProperties.setProperty("variables", Utilities.jsonStringify(variables));
    Utilities.sleep(500);
    ScriptProperties.setProperty("status", Utilities.jsonStringify(status));
    var currentTriggers = ScriptApp.getScriptTriggers();
    for(i in currentTriggers){
      ScriptApp.deleteTrigger(currentTriggers[i]);
    }
    ScriptApp.newTrigger('activityReport').timeBased().everyMinutes(5).create();
}