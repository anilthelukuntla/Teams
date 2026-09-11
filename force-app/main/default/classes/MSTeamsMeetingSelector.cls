public with sharing class MSTeamsMeetingSelector {
    public static List<MSTeamsMeetingSummary> getMeetings(Id recordId) {
        List<MSTeamsMeetingSummary> results = new List<MSTeamsMeetingSummary>();
        Set<Id> meetingIds = MSTeamsRecordLinkService.getMeetingIds(recordId);
        if (meetingIds.isEmpty()) {
            return results;
        }
        for (MS_Team_Call_Schedule__c row : [
            SELECT Id, Subject__c, Body_HTML__c, Meeting_Type__c,
                Status__c, Attendee_Emails__c,
                Start_Date_Time__c, End_Date_Time__c, Microsoft_Time_Zone__c,
                Organizer_Email__c, Web_Link__c, Teams_Join_URL__c,
                Recording_Requested__c, Transcription_Requested__c,
                Last_Synchronized_At__c, Last_Graph_Error__c
            FROM MS_Team_Call_Schedule__c
            WHERE Id IN :meetingIds
            WITH USER_MODE
            ORDER BY Start_Date_Time__c DESC
            LIMIT 200
        ]) {
            results.add(toSummary(row));
        }
        return results;
    }

    public static MSTeamsMeetingDetails getMeetingDetails(
        Id recordId,
        Id scheduleId
    ) {
        MSTeamsRecordLinkService.assertMeetingLinked(recordId, scheduleId);
        MS_Team_Call_Schedule__c schedule = [
            SELECT Id, Subject__c, Body_HTML__c, Meeting_Type__c,
                Status__c, Attendee_Emails__c,
                Start_Date_Time__c, End_Date_Time__c, Microsoft_Time_Zone__c,
                Organizer_Email__c, Web_Link__c, Teams_Join_URL__c,
                Recording_Requested__c, Transcription_Requested__c,
                Last_Synchronized_At__c, Last_Graph_Error__c,
                Meeting_Started_At__c, Meeting_Ended_At__c,
                Actual_Duration_Seconds__c
            FROM MS_Team_Call_Schedule__c
            WHERE Id = :scheduleId
            WITH USER_MODE
            LIMIT 1
        ];
        MSTeamsMeetingDetails result = new MSTeamsMeetingDetails();
        result.meeting = toSummary(schedule);
        result.startedAt = schedule.Meeting_Started_At__c;
        result.endedAt = schedule.Meeting_Ended_At__c;
        result.actualDurationSeconds = schedule.Actual_Duration_Seconds__c;
        result.participants = new List<MSTeamsParticipant>();
        result.artifacts = new List<MSTeamsArtifact>();

        Map<Id, MSTeamsParticipant> participantsById =
            new Map<Id, MSTeamsParticipant>();
        for (MS_Team_Meeting_Participant__c row : [
            SELECT Id, Display_Name__c, Email__c, Role__c, Attendee_Type__c,
                Invitation_Response__c, Invitation_Response_At__c,
                Actually_Joined__c, First_Joined_At__c, Last_Left_At__c,
                Total_Attendance_Seconds__c
            FROM MS_Team_Meeting_Participant__c
            WHERE Meeting_Schedule__c = :scheduleId
            WITH USER_MODE
            ORDER BY Display_Name__c, Email__c
        ]) {
            MSTeamsParticipant participant = new MSTeamsParticipant();
            participant.id = row.Id;
            participant.displayName = row.Display_Name__c;
            participant.email = row.Email__c;
            participant.role = row.Role__c;
            participant.attendeeType = row.Attendee_Type__c;
            participant.invitationResponse = row.Invitation_Response__c;
            participant.invitationResponseAt = row.Invitation_Response_At__c;
            participant.actuallyJoined = row.Actually_Joined__c;
            participant.firstJoinedAt = row.First_Joined_At__c;
            participant.lastLeftAt = row.Last_Left_At__c;
            participant.totalAttendanceSeconds = row.Total_Attendance_Seconds__c;
            participant.intervals = new List<MSTeamsAttendanceInterval>();
            participantsById.put(row.Id, participant);
            result.participants.add(participant);
        }
        if (!participantsById.isEmpty()) {
            for (MS_Team_Attendance_Interval__c row : [
                SELECT Id, Participant__c, Joined_At__c, Left_At__c,
                    Duration_Seconds__c
                FROM MS_Team_Attendance_Interval__c
                WHERE Participant__c IN :participantsById.keySet()
                WITH USER_MODE
                ORDER BY Joined_At__c
            ]) {
                MSTeamsAttendanceInterval interval =
                    new MSTeamsAttendanceInterval();
                interval.id = row.Id;
                interval.joinedAt = row.Joined_At__c;
                interval.leftAt = row.Left_At__c;
                interval.durationSeconds = row.Duration_Seconds__c;
                participantsById.get(row.Participant__c).intervals.add(interval);
            }
        }
        for (MS_Team_Meeting_Artifact__c row : [
            SELECT Id, Artifact_Type__c, Processing_Status__c, Created_At__c,
                Ended_At__c, Content_URL__c, Last_Error__c
            FROM MS_Team_Meeting_Artifact__c
            WHERE Meeting_Schedule__c = :scheduleId
            WITH USER_MODE
            ORDER BY Created_At__c DESC
        ]) {
            MSTeamsArtifact artifact = new MSTeamsArtifact();
            artifact.id = row.Id;
            artifact.artifactType = row.Artifact_Type__c;
            artifact.processingStatus = row.Processing_Status__c;
            artifact.createdAt = row.Created_At__c;
            artifact.endedAt = row.Ended_At__c;
            artifact.contentUrl = row.Content_URL__c;
            artifact.lastError = row.Last_Error__c;
            result.artifacts.add(artifact);
        }
        return result;
    }

    private static MSTeamsMeetingSummary toSummary(
        MS_Team_Call_Schedule__c row
    ) {
        MSTeamsMeetingSummary result = new MSTeamsMeetingSummary();
        result.id = row.Id;
        result.subject = row.Subject__c;
        result.description = row.Body_HTML__c;
        result.meetingType = row.Meeting_Type__c;
        result.status = row.Status__c;
        result.attendeeEmails = row.Attendee_Emails__c;
        result.startDateTime = row.Start_Date_Time__c;
        result.endDateTime = row.End_Date_Time__c;
        result.timeZone = row.Microsoft_Time_Zone__c;
        result.organizerEmail = row.Organizer_Email__c;
        result.webLink = row.Web_Link__c;
        result.joinUrl = row.Teams_Join_URL__c;
        result.recordingRequested = row.Recording_Requested__c;
        result.transcriptionRequested = row.Transcription_Requested__c;
        result.lastSynchronizedAt = row.Last_Synchronized_At__c;
        result.lastGraphError = row.Last_Graph_Error__c;
        return result;
    }
}
