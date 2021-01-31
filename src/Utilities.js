function onInstall(){
  onOpen();
}

function onOpen() {  
  ss.addMenu("Forum Statistics", [{name: "Get a Report", functionName: "uiApp_"}]);
}

function daysInMonth_(month, year) {
    return 32 - new Date(year, month, 32).getDate();
}

function countSendsPerDaysOfWeek_(variables, date) {
    switch (Utilities.formatDate(date, variables.userTimeZone, "EEEE")) {
    case "Monday":
        variables.dayOfWeek[0]++;
        break;
    case "Tuesday":
        variables.dayOfWeek[1]++;
        break;
    case "Wednesday":
        variables.dayOfWeek[2]++;
        break;
    case "Thursday":
        variables.dayOfWeek[3]++;
        break;
    case "Friday":
        variables.dayOfWeek[4]++;
        break;
    case "Saturday":
        variables.dayOfWeek[5]++;
        break;
    case "Sunday":
        variables.dayOfWeek[6]++;
        break;
    }
    return variables;
}

function calcMessagesLength_(variables, body) {
    var eom1 = body.indexOf("<pre>");
    var eom2 = body.indexOf("<div class=\"gmail_quote");
    var eom = Math.min(eom1, eom2);
    if (eom == -1) {
        eom = Math.max(eom1, eom2);
    }
    if (eom != -1) {
        body = body.substring(0, eom);
    }
    var matches = body.replace(/<[^<|>]+?>|&nbsp;/gi, ' ').match(/\b/g);
    var count = 0;
    if (matches) {
        count = matches.length / 2;
    }
    if (count < 10) {
        variables.messagesLength[0]++;
    }
    else if (count < 30) {
        variables.messagesLength[1]++;
    }
    else if (count < 50) {
        variables.messagesLength[2]++;
    }
    else if (count < 100) {
        variables.messagesLength[3]++;
    }
    else if (count < 200) {
        variables.messagesLength[4]++;
    }
    else {
        variables.messagesLength[5]++;
    }
    return [variables, count];
}

function calcWaitingTime_(variables, date, timeOfFirstMessage, youStartedTheConversation) {
    var waitingTime = Math.round((date.getTime() - timeOfFirstMessage) / 60000);
    if (waitingTime < 15) {
        variables.timeBeforeFirstResponse[0]++;
    }
    else if (waitingTime < 60) {
        variables.timeBeforeFirstResponse[1]++;
    }
    else if (waitingTime < 240) {
        variables.timeBeforeFirstResponse[2]++;
    }
    else if (waitingTime < 1440) {
        variables.timeBeforeFirstResponse[3]++;
    }
    else if (waitingTime < 2880) {
        variables.timeBeforeFirstResponse[4]++;
    }
    else {
        variables.timeBeforeFirstResponse[5]++;
    }
    return variables;
}