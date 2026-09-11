public class MSTeamsMeetingDetails {
    @AuraEnabled public MSTeamsMeetingSummary meeting { get; set; }
    @AuraEnabled public Datetime startedAt { get; set; }
    @AuraEnabled public Datetime endedAt { get; set; }
    @AuraEnabled public Decimal actualDurationSeconds { get; set; }
    @AuraEnabled public List<MSTeamsParticipant> participants { get; set; }
    @AuraEnabled public List<MSTeamsArtifact> artifacts { get; set; }
    @AuraEnabled public List<MSTeamsRelatedRecord> relatedRecords { get; set; }
    @AuraEnabled public String location { get; set; }
    @AuraEnabled public String meetingAccess { get; set; }
    @AuraEnabled public Boolean allowMeetingChat { get; set; }
    @AuraEnabled public Boolean allowScreenSharing { get; set; }
    @AuraEnabled public Boolean lobbyRequired { get; set; }
}
