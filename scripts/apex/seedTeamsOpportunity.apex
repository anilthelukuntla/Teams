// Synthetic Salesforce-only fixtures. No Graph calls, invitations, or real artifact URLs.
// Re-running stops if the fixture Opportunity already exists.
String fixture = 'TEAMS-UI-TEST-20260916';
String opportunityName = 'Teams UI Test - 10 Meeting Scenarios';
System.assertEquals(true, [SELECT IsSandbox FROM Organization LIMIT 1].IsSandbox, 'Sandbox only.');
System.assertEquals(0, [SELECT COUNT() FROM Opportunity WHERE Name = :opportunityName], 'Fixture already exists; do not duplicate.');
Savepoint checkpoint = Database.setSavepoint();
try {
    Opportunity opp = new Opportunity(Name=opportunityName, StageName='Qualification', CloseDate=Date.today().addDays(60), Amount=125000,
        Description=fixture + ': Synthetic Teams UI testing. No real Microsoft meetings or invitations.');
    insert opp;
    List<String> statuses = new List<String>{'Scheduled','Scheduled','Scheduled','Scheduled','In Progress','Completed','Completed','Cancelled','Deleted','Sync Failed'};
    List<String> subjects = new List<String>{'Discovery Call','Product Demonstration','Pricing Discussion','Contract Review','Solution Workshop','Requirements Review','Technical Review','Rescheduled Planning Session','Removed Duplicate Meeting','Sync Recovery Scenario'};
    List<MS_Team_Call_Schedule__c> meetings = new List<MS_Team_Call_Schedule__c>();
    List<Datetime> starts = new List<Datetime>();
    for (Integer i=0; i<10; i++) {
        Datetime startAt = i<4 ? System.now().addDays(i+1) : (i==4 ? System.now().addMinutes(-10) : System.now().addDays(-(i-3)));
        starts.add(startAt);
        MS_Team_Call_Schedule__c meeting = new MS_Team_Call_Schedule__c(
            Subject__c='[TEST] '+subjects[i], Status__c=statuses[i], Meeting_Type__c=i==1?'Product Demo':'Customer Meeting',
            Start_Date_Time__c=startAt.format('yyyy-MM-dd\'T\'HH:mm:ss','Asia/Kolkata'),
            End_Date_Time__c=startAt.addMinutes(60).format('yyyy-MM-dd\'T\'HH:mm:ss','Asia/Kolkata'), Microsoft_Time_Zone__c='India Standard Time',
            Body_HTML__c='<p>Synthetic UI test: '+subjects[i]+'. Review requirements, decisions, and next steps.</p>',
            Organizer_Email__c='organizer@example.invalid', Attendee_Emails__c='alex@example.invalid;priya@example.invalid;sam@example.invalid',
            Location__c='Teams UI preview (synthetic)', Meeting_Access__c='Invited People',
            Allow_Meeting_Chat__c=true, Allow_Screen_Sharing__c=true, Allow_New_Time_Proposals__c=true,
            Lobby_Required__c=i==3, Recording_Requested__c=i!=2, Transcription_Requested__c=i!=2,
            Reminder_Minutes_Before_Start__c=15, Teams_Meeting_Requested__c=true, Is_Online_Meeting__c=true,
            Graph_Transaction_ID__c=fixture+'-'+i, Last_Synchronized_At__c=System.now());
        if(i==4 || i==5 || i==6) meeting.Meeting_Started_At__c=startAt;
        if(i==5 || i==6) { meeting.Meeting_Ended_At__c=startAt.addMinutes(60); meeting.Actual_Duration_Seconds__c=3600; }
        if(i==7) { meeting.Cancelled_At__c=System.now(); meeting.Cancelled_By__c=UserInfo.getUserId(); meeting.Cancellation_Reason__c='Synthetic cancellation example'; }
        if(i==8) { meeting.Deleted_At__c=System.now(); meeting.Deleted_By__c=UserInfo.getUserId(); meeting.Deletion_Justification__c='Synthetic soft-delete example'; }
        if(i==9) meeting.Last_Graph_Error__c='Synthetic error: Microsoft connection requires reauthentication.';
        meetings.add(meeting);
    }
    insert meetings;
    List<MS_Teams_Record_Link__c> links = new List<MS_Teams_Record_Link__c>();
    List<MS_Team_Meeting_Participant__c> people = new List<MS_Team_Meeting_Participant__c>();
    List<MS_Graph_Subscription__c> subscriptions = new List<MS_Graph_Subscription__c>();
    List<MS_Teams_Chat_Message__c> chats = new List<MS_Teams_Chat_Message__c>();
    List<MS_Team_Meeting_Artifact__c> artifacts = new List<MS_Team_Meeting_Artifact__c>();
    List<String> names = new List<String>{'Alex Morgan','Priya Shah','Sam Lee'};
    List<String> emails = new List<String>{'alex@example.invalid','priya@example.invalid','sam@example.invalid'};
    for (Integer i=0; i<10; i++) {
        Id meetingId=meetings[i].Id;
        links.add(new MS_Teams_Record_Link__c(Meeting_Schedule__c=meetingId, External_Key__c=fixture+'-link-'+i,
            Record_ID__c=opp.Id, Object_API_Name__c='Opportunity', Record_Label__c=opp.Name, Relationship_Role__c='Primary'));
        for (Integer j=0; j<3; j++) {
            Boolean joined=(i==4 || i==5 || i==6) && j<2;
            people.add(new MS_Team_Meeting_Participant__c(Meeting_Schedule__c=meetingId, External_Key__c=fixture+'-person-'+i+'-'+j,
                Display_Name__c=names[j], Email__c=emails[j], Role__c=j==0?'Presenter':'Attendee', Attendee_Type__c=j==2?'Optional':'Required',
                Invitation_Response__c=j==0?'Accepted':(j==1?'Tentatively Accepted':'Not Responded'), Invitation_Response_At__c=starts[i].addDays(-1),
                Actually_Joined__c=joined, First_Joined_At__c=joined?starts[i]:null,
                Last_Left_At__c=joined && i!=4?starts[i].addMinutes(55):null, Total_Attendance_Seconds__c=joined?(i==4?600:3300):0));
        }
        subscriptions.add(new MS_Graph_Subscription__c(Meeting_Schedule__c=meetingId, Graph_Subscription_ID__c=fixture+'-subscription-'+i,
            Client_State_Hash__c=EncodingUtil.convertToHex(Crypto.generateDigest('SHA-256',Blob.valueOf(fixture+'-'+i))),
            Status__c='Expired', Expiration__c=System.now().addDays(-1), Change_Type__c='updated',
            Resource__c='Synthetic UI fixture; no Microsoft subscription', Last_Error__c='Test-only expired subscription. Not registered with Microsoft.'));
        if(i==4 || i==5 || i==6) {
            for(Integer j=0; j<3; j++) chats.add(new MS_Teams_Chat_Message__c(Meeting_Schedule__c=meetingId,
                External_Key__c=fixture+'-chat-'+i+'-'+j, Graph_Chat_ID__c=fixture+'-thread-'+i, Graph_Message_ID__c=fixture+'-message-'+i+'-'+j,
                Sender_Display_Name__c=names[j], Sender_Email__c=emails[j], Body_Content_Type__c=j==1?'HTML':'Text',
                Body__c=j==1?'<p>We will share the action items after this test meeting.</p>':'Sample meeting chat for UI testing.',
                Message_Type__c='message', Sent_At__c=starts[i].addMinutes(j+1), Last_Synchronized_At__c=System.now()));
        }
        if(i==5 || i==6) for(String kind:new List<String>{'Recording','Transcript'}) artifacts.add(new MS_Team_Meeting_Artifact__c(
            Meeting_Schedule__c=meetingId, Artifact_Type__c=kind, External_Key__c=fixture+'-artifact-'+i+'-'+kind,
            Graph_Artifact_ID__c=fixture+'-artifact-'+i+'-'+kind, Created_At__c=starts[i], Ended_At__c=starts[i].addMinutes(60),
            Processing_Status__c=i==5?'Completed':'Failed', Last_Error__c=i==5?'Synthetic metadata only; no downloadable file.':'Synthetic artifact processing failure.'));
    }
    insert links; insert people; insert subscriptions; insert chats; insert artifacts;
    List<MS_Team_Attendance_Interval__c> intervals=new List<MS_Team_Attendance_Interval__c>();
    for(Integer k=0;k<people.size();k++) if(people[k].Actually_Joined__c) intervals.add(new MS_Team_Attendance_Interval__c(
        Participant__c=people[k].Id, External_Key__c=fixture+'-attendance-'+k, Attendance_Report_ID__c=fixture+'-report-'+k,
        Joined_At__c=people[k].First_Joined_At__c, Left_At__c=people[k].Last_Left_At__c, Duration_Seconds__c=people[k].Total_Attendance_Seconds__c));
    insert intervals;
    List<MS_Teams_Sync_Log__c> logs=new List<MS_Teams_Sync_Log__c>();
    List<String> operations=new List<String>{'Create Meeting','Update Meeting','Cancel Meeting','Delete Meeting','Sync Meeting','Sync Attendance','Sync Artifacts','Sync Chat','Create Subscription','Renew Subscription','Process Notification'};
    List<String> logStates=new List<String>{'Queued','Running','Succeeded','Failed','Retry Scheduled','Skipped'};
    for(Integer i=0;i<operations.size();i++) logs.add(new MS_Teams_Sync_Log__c(Meeting_Schedule__c=meetings[Math.mod(i,10)].Id,
        Subscription__c=subscriptions[Math.mod(i,10)].Id, Operation__c=operations[i], Status__c=logStates[Math.mod(i,6)],
        Attempt_Number__c=1, Started_At__c=System.now().addMinutes(-5), Correlation_ID__c=fixture+'-log-'+i,
        Records_Processed__c=0, Error_Message__c='Synthetic UI log; no job has been queued or executed.'));
    insert logs;
    // Reuse current settings; only create a disabled user-specific fixture if none exists.
    MS_Teams_Setting__c settings=MS_Teams_Setting__c.getInstance(UserInfo.getUserId());
    if(settings.Id==null) {
        settings=new MS_Teams_Setting__c(SetupOwnerId=UserInfo.getUserId(), Sync_Attendance__c=false, Sync_Meetings__c=false,
            Sync_Recordings__c=false, Sync_Transcripts__c=false, Sync_Window_Days__c=7);
        insert settings;
    }
    System.debug('TEAMS_FIXTURE_RESULT '+JSON.serialize(new Map<String,Object>{'opportunityId'=>opp.Id,'opportunityName'=>opp.Name,
        'meetings'=>meetings.size(),'recordLinks'=>links.size(),'participants'=>people.size(),'attendanceIntervals'=>intervals.size(),
        'artifacts'=>artifacts.size(),'chatMessages'=>chats.size(),'subscriptions'=>subscriptions.size(),'syncLogs'=>logs.size(),'settingsId'=>settings.Id}));
} catch(Exception e) {
    Database.rollback(checkpoint);
    System.debug(LoggingLevel.ERROR,'TEAMS_FIXTURE_ROLLED_BACK '+e.getMessage()+' at line '+e.getLineNumber());
    throw e;
}
