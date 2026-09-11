public class MSTeamsActionRequest {
    @AuraEnabled public Id recordId { get; set; }
    @AuraEnabled public Id scheduleId { get; set; }
    @AuraEnabled public String reason { get; set; }
}
