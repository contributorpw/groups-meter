function sendReport_(variables) {
    var status = Utilities.jsonParse(ScriptProperties.getProperty("status"));
    var report = "<h2 style=\"color:#cccccc; font-family:trebuchet ms;\">Forum statistics - ";
    if (status.customReport) {
        report += "From " + Utilities.formatDate(new Date(variables.startDate), variables.userTimeZone, 'MM/dd/yyyy');
        report += " to " + Utilities.formatDate(new Date(variables.endDate), variables.userTimeZone, 'MM/dd/yyyy') + "</h2>";
    }
    else {
        report += Utilities.formatDate(new Date(variables.previous), variables.userTimeZone, "MMMM") + "</h2>";
    }
    report += "<p><table style=\"border-collapse: collapse;\"><tr><td style=\"border: 0px solid white; width: 150px; padding-left: 5px; color: white; background-color: #009754;\"><h3>";
    report += variables.nbrOfConversations + " active topics</h3></td>";
    report += "<td style=\"border: 0px solid white; width: 50px;\"></td>";
    report += "<td style=\"border: 0px solid white; margin-left: 20px; width: 150px; padding-left: 5px; color: white; background-color: #5fb5d8;\"><h3>";
    report += variables.nbrOfMessages+" replies</h3></td></tr></table></p>";
    report += "<p>"+Math.round((variables.nbrOfConversations-variables.nbrOfEmailsPerConversation[1]) * 10000 / (variables.nbrOfConversations)) / 100+"% of those topics got an answer.</p>";
  
    report += "<h3>Daily Traffic</h3><img src=\'";

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'Time');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Received');

    var time = '';
    for (var i = 0; i < variables.timeOfEmailsReceived.length; i++) { //create the rows
        switch (i) {
        case 0:
            time = 'Midnight';
            break;
        case 6:
            time = '6 AM';
            break;
        case 12:
            time = 'NOON';
            break;
        case 18:
            time = '6 PM';
            break;
        default:
            time = '';
            break;
        }
        dataTable.addRow([time, variables.timeOfEmailsReceived[i]]);
    }
    dataTable.build();
    var chartAverageFlow = Charts.newAreaChart().setDataTable(dataTable).setLegendPosition(Charts.Position.NONE).setDimensions(650, 400).build();

    report += "cid:Averageflow\'/><h3>Monthly Traffic</h3><img src=\'";

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'Date');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Received');

    for (var i = 0; i < variables.dayOfEmailsReceived.length; i++) { //create the rows
        dataTable.addRow([(i + 1).toString(), variables.dayOfEmailsReceived[i]]);
    }
    dataTable.build();
    var chartDate = Charts.newAreaChart().setDataTable(dataTable).setLegendPosition(Charts.Position.NONE).setDimensions(650, 400).build();

    report += "cid:Date\'/><h3>Weekly Traffic</h3><img src=\'";

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'dayOfWeek');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Received');
    var dayTags = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    for (var i = 0; i < dayTags.length; i++) {
        dataTable.addRow([dayTags[i], variables.dayOfWeek[i] * 100 / variables.nbrOfMessages]);
    }
    dataTable.build();
    var chartDayOfWeek = Charts.newColumnChart().setDataTable(dataTable).setLegendPosition(Charts.Position.NONE).setYAxisTitle("in % of messages").setDimensions(650, 400).build();

    report += "cid:DayOfWeek\'/>";

    if (status.customReport) {
        report += "<h3>Month by month</h3><img src=\'";

        var dataTable = Charts.newDataTable();
        dataTable.addColumn(Charts.ColumnType['STRING'], 'Month');
        dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Received');
        var monthTags = ['J', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D'];
        for (var i = 0; i < monthTags.length; i++) {
            dataTable.addRow([monthTags[i], variables.monthOfEmailsReceived[i] * 100 / variables.nbrOfMessages]);
        }
        dataTable.build();
        var chartMonths = Charts.newColumnChart().setDataTable(dataTable).setLegendPosition(Charts.Position.NONE).setYAxisTitle("in % of messages").setDimensions(650, 400).build();

        report += "cid:Months\'/>";
    }

    report += "<h3>Thread Lengths</h3><img src=\'";

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Length');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'Number of conversations');

    for (var i = 0; i < variables.nbrOfEmailsPerConversation.length; i++) { //create the rows
        dataTable.addRow([i, variables.nbrOfEmailsPerConversation[i]]);
    }
    dataTable.build();
    var chartConversation = Charts.newScatterChart().setDataTable(dataTable).setDimensions(650, 400).setXAxisLogScale().setYAxisLogScale().setLegendPosition(Charts.Position.NONE).build();

    report += "cid:Conversation\'/><h3>Top threads</h3>";
    report += "<table style=\"border-collapse: collapse; border: 0px solid white;\">";
    for (var i = 0; i < variables.topThreads.length; i++) {
        report += "<tr><td>" + variables.topThreads[i][0] + "</td><td>" + variables.topThreads[i][1] + "</td></tr>";
    }
    report += "</table><br><h3>Time Before First Response</h3><img src=\'";
    var count = 0;
    for (var k = 0; k < 6; k++) {
        count += variables.timeBeforeFirstResponse[k];
    }
    for (var k = 0; k < 6; k++) {
        if (variables.timeBeforeFirstResponse[k] != 0) {
            variables.timeBeforeFirstResponse[k] = variables.timeBeforeFirstResponse[k] * 100 / count;
        }
    }
    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'Time');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'When people answer');
    dataTable.addRow(['< 15min', variables.timeBeforeFirstResponse[0]]);
    dataTable.addRow(['< 1hr', variables.timeBeforeFirstResponse[1]]);
    dataTable.addRow(['< 4hrs', variables.timeBeforeFirstResponse[2]]);
    dataTable.addRow(['< 1day', variables.timeBeforeFirstResponse[3]]);
    dataTable.addRow(['< 2days', variables.timeBeforeFirstResponse[4]]);
    dataTable.addRow(['More', variables.timeBeforeFirstResponse[5]]);

    dataTable.build();
    var chartWaitingTime = Charts.newColumnChart().setDataTable(dataTable).setLegendPosition(Charts.Position.NONE).setYAxisTitle("in % of messages").setDimensions(650, 400).build();
    
    var lessThan2Days = variables.timeBeforeFirstResponse[0] + variables.timeBeforeFirstResponse[1] + variables.timeBeforeFirstResponse[2];
    lessThan2Days+= variables.timeBeforeFirstResponse[3] + variables.timeBeforeFirstResponse[4];
    lessThan2Days = Math.round(lessThan2Days * 100) / 100;
    
  report += "cid:WaitingTime\'/><p>"+lessThan2Days+"% got a first response within 2 days.</p><h3>Word Count</h3><img src=\'";

    count = 0;
    for (var k = 0; k < 6; k++) {
        count += variables.messagesLength[k];
    }
    for (var k = 0; k < 6; k++) {
        if (variables.messagesLength[k] != 0) {
            variables.messagesLength[k] = variables.messagesLength[k] * 100 / count;
        }
    }

    var dataTable = Charts.newDataTable();
    dataTable.addColumn(Charts.ColumnType['STRING'], 'Words');
    dataTable.addColumn(Charts.ColumnType['NUMBER'], 'When people answer');
    dataTable.addRow(['< 10 words', variables.messagesLength[0]]);
    dataTable.addRow(['< 30', variables.messagesLength[1]]);
    dataTable.addRow(['< 50', variables.messagesLength[2]]);
    dataTable.addRow(['< 100', variables.messagesLength[3]]);
    dataTable.addRow(['< 200', variables.messagesLength[4]]);
    dataTable.addRow(['More', variables.messagesLength[5]]);

    dataTable.build();
    var chartMessagesLength = Charts.newColumnChart().setDataTable(dataTable).setYAxisTitle("in % of messages").setLegendPosition(Charts.Position.NONE).setDimensions(650, 400).build();
    report += "cid:MessagesLength\'/>";
      
    var inlineImages = {};
    inlineImages['Averageflow'] = chartAverageFlow;
    inlineImages['Date'] = chartDate;
    inlineImages['DayOfWeek'] = chartDayOfWeek;
    if (status.customReport) {
        inlineImages['Months'] = chartMonths;
    }
    inlineImages['Conversation'] = chartConversation;
    inlineImages['WaitingTime'] = chartWaitingTime;
    inlineImages['MessagesLength'] = chartMessagesLength;


    MailApp.sendEmail(variables.user, "Forum statistics", report, {
        htmlBody: report,
        inlineImages: inlineImages
    });
    status.reportSent = "yes";
    ScriptProperties.setProperty("status", Utilities.jsonStringify(status));
    var currentTriggers = ScriptApp.getScriptTriggers();
    for (i in currentTriggers) {
      ScriptApp.deleteTrigger(currentTriggers[i]);
    }
    ScriptApp.newTrigger('activityReport').timeBased().everyDays(1).atHour(1).create();
}