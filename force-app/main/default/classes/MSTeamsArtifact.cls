public class MSTeamsArtifact {
    @AuraEnabled public Id id { get; set; }
    @AuraEnabled public String artifactType { get; set; }
    @AuraEnabled public String processingStatus { get; set; }
    @AuraEnabled public Datetime createdAt { get; set; }
    @AuraEnabled public Datetime endedAt { get; set; }
    @AuraEnabled public String contentUrl { get; set; }
    @AuraEnabled public String lastError { get; set; }
}
