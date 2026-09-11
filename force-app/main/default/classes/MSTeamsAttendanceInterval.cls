public class MSTeamsAttendanceInterval {
    @AuraEnabled public Id id { get; set; }
    @AuraEnabled public Datetime joinedAt { get; set; }
    @AuraEnabled public Datetime leftAt { get; set; }
    @AuraEnabled public Decimal durationSeconds { get; set; }
}
