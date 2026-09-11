public class MSTeamsMeetingRequest {
    @AuraEnabled public Id recordId { get; set; }
    @AuraEnabled public Id scheduleId { get; set; }
    @AuraEnabled public String subject { get; set; }
    @AuraEnabled public String description { get; set; }
    @AuraEnabled public List<String> attendeeEmails { get; set; }
    @AuraEnabled public String startDateTime { get; set; }
    @AuraEnabled public String endDateTime { get; set; }
    @AuraEnabled public String microsoftTimeZone { get; set; }
    @AuraEnabled public Boolean recordAutomatically { get; set; }
    @AuraEnabled public Boolean enableTranscription { get; set; }
    @AuraEnabled public String meetingType { get; set; }
    @AuraEnabled public Boolean allowMeetingChat { get; set; }
    @AuraEnabled public Boolean allowScreenSharing { get; set; }
    @AuraEnabled public Boolean lobbyRequired { get; set; }
    @AuraEnabled public String meetingAccess { get; set; }
    @AuraEnabled public String teamsChannelId { get; set; }
    @AuraEnabled public String location { get; set; }
    @AuraEnabled public Integer reminderMinutesBeforeStart { get; set; }
    @AuraEnabled public Boolean allowNewTimeProposals { get; set; }
}
