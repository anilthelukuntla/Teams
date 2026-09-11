public class MSTeamsParticipant {
    @AuraEnabled public Id id { get; set; }
    @AuraEnabled public String displayName { get; set; }
    @AuraEnabled public String email { get; set; }
    @AuraEnabled public String role { get; set; }
    @AuraEnabled public String attendeeType { get; set; }
    @AuraEnabled public String invitationResponse { get; set; }
    @AuraEnabled public Datetime invitationResponseAt { get; set; }
    @AuraEnabled public Boolean actuallyJoined { get; set; }
    @AuraEnabled public Datetime firstJoinedAt { get; set; }
    @AuraEnabled public Datetime lastLeftAt { get; set; }
    @AuraEnabled public Decimal totalAttendanceSeconds { get; set; }
    @AuraEnabled public List<MSTeamsAttendanceInterval> intervals { get; set; }
}
