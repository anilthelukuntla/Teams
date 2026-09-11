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
}
