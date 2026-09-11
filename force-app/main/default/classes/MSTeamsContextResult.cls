public class MSTeamsContextResult {
    @AuraEnabled public Id recordId { get; set; }
    @AuraEnabled public String objectApiName { get; set; }
    @AuraEnabled public String recordName { get; set; }
    @AuraEnabled public List<String> defaultRecipientEmails { get; set; }
    @AuraEnabled public String defaultSubject { get; set; }
    @AuraEnabled public String defaultDescription { get; set; }
}
