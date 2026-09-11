public class MSTeamsMeetingSummary {
    @AuraEnabled public Id id { get; set; }
    @AuraEnabled public String subject { get; set; }
    @AuraEnabled public String description { get; set; }
    @AuraEnabled public String meetingType { get; set; }
    @AuraEnabled public String status { get; set; }
    @AuraEnabled public String attendeeEmails { get; set; }
    @AuraEnabled public String startDateTime { get; set; }
    @AuraEnabled public String endDateTime { get; set; }
    @AuraEnabled public String timeZone { get; set; }
    @AuraEnabled public String organizerEmail { get; set; }
    @AuraEnabled public String webLink { get; set; }
    @AuraEnabled public String joinUrl { get; set; }
    @AuraEnabled public Boolean recordingRequested { get; set; }
    @AuraEnabled public Boolean transcriptionRequested { get; set; }
    @AuraEnabled public Datetime lastSynchronizedAt { get; set; }
    @AuraEnabled public String lastGraphError { get; set; }
}
