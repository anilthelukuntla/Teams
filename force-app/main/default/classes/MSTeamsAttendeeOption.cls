public class MSTeamsAttendeeOption {
    @AuraEnabled public Id recordId { get; set; }
    @AuraEnabled public String objectApiName { get; set; }
    @AuraEnabled public String displayName { get; set; }
    @AuraEnabled public String title { get; set; }
    @AuraEnabled public String email { get; set; }
    @AuraEnabled public String sourceLabel { get; set; }
}
