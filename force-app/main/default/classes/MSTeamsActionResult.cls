public class MSTeamsActionResult {
    @AuraEnabled public Boolean success { get; set; }
    @AuraEnabled public String message { get; set; }
    @AuraEnabled public Id scheduleId { get; set; }
    @AuraEnabled public String eventId { get; set; }
    @AuraEnabled public String joinUrl { get; set; }
    @AuraEnabled public String webLink { get; set; }

    public static MSTeamsActionResult succeeded(Id scheduleId, String message) {
        MSTeamsActionResult result = new MSTeamsActionResult();
        result.success = true;
        result.scheduleId = scheduleId;
        result.message = message;
        return result;
    }
}
