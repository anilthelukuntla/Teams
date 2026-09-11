public class MSTeamsConnectionResult {
    @AuraEnabled public Boolean connected { get; set; }
    @AuraEnabled public Integer statusCode { get; set; }
    @AuraEnabled public String status { get; set; }
    @AuraEnabled public String message { get; set; }
    @AuraEnabled public String microsoftUserId { get; set; }
    @AuraEnabled public String displayName { get; set; }
    @AuraEnabled public String userPrincipalName { get; set; }
    @AuraEnabled public String email { get; set; }
}
