public class MSTeamsRelatedRecord {
    @AuraEnabled public String recordId { get; set; }
    @AuraEnabled public String objectApiName { get; set; }
    @AuraEnabled public String recordLabel { get; set; }
    @AuraEnabled public String relationshipRole { get; set; }
}
