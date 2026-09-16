Opportunity opp=[SELECT Id,Name FROM Opportunity WHERE Name='Teams UI Test - 10 Meeting Scenarios' LIMIT 1];
Set<Id> ids=new Set<Id>();
for(MS_Teams_Record_Link__c link:[SELECT Meeting_Schedule__c FROM MS_Teams_Record_Link__c WHERE Record_ID__c=:opp.Id]) ids.add(link.Meeting_Schedule__c);
System.assertEquals(10,ids.size(),'Expected exactly 10 linked meetings.');
Map<String,Object> counts=new Map<String,Object>{
    'opportunityId'=>opp.Id,
    'meetings'=>[SELECT COUNT() FROM MS_Team_Call_Schedule__c WHERE Id IN :ids],
    'participants'=>[SELECT COUNT() FROM MS_Team_Meeting_Participant__c WHERE Meeting_Schedule__c IN :ids],
    'attendanceIntervals'=>[SELECT COUNT() FROM MS_Team_Attendance_Interval__c WHERE Participant__r.Meeting_Schedule__c IN :ids],
    'artifacts'=>[SELECT COUNT() FROM MS_Team_Meeting_Artifact__c WHERE Meeting_Schedule__c IN :ids],
    'chatMessages'=>[SELECT COUNT() FROM MS_Teams_Chat_Message__c WHERE Meeting_Schedule__c IN :ids],
    'subscriptions'=>[SELECT COUNT() FROM MS_Graph_Subscription__c WHERE Meeting_Schedule__c IN :ids],
    'syncLogs'=>[SELECT COUNT() FROM MS_Teams_Sync_Log__c WHERE Meeting_Schedule__c IN :ids]
};
System.debug('TEAMS_VERIFIED '+JSON.serialize(counts));
System.debug('TEAMS_STATUSES '+JSON.serialize([SELECT Status__c,COUNT(Id) total FROM MS_Team_Call_Schedule__c WHERE Id IN :ids GROUP BY Status__c]));
