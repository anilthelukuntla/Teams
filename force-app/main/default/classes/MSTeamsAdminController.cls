public with sharing class MSTeamsAdminController {
    @AuraEnabled(cacheable=true)
    public static SetupInfo getSetupInfo() {
        Organization orgInfo = [SELECT Id, Name FROM Organization LIMIT 1];
        MS_Teams_Setting__c settings = MS_Teams_Setting__c.getOrgDefaults();
        AggregateResult latestSync = [
            SELECT MAX(Completed_At__c) completedAt
            FROM MS_Teams_Sync_Log__c
            WHERE Status__c = 'Succeeded'
        ];

        SetupInfo result = new SetupInfo();
        result.tenantName = orgInfo.Name;
        result.tenantId = String.valueOf(orgInfo.Id);
        result.connectedUsers = [
            SELECT COUNT()
            FROM PermissionSetAssignment
            WHERE PermissionSet.Name = 'MSTeams_Meeting_Organizer'
        ];
        result.lastSync = (Datetime) latestSync.get('completedAt');
        result.activeSubscriptions = [
            SELECT COUNT()
            FROM MS_Graph_Subscription__c
            WHERE Status__c = 'Active'
        ];
        result.failedSyncs = [
            SELECT COUNT()
            FROM MS_Teams_Sync_Log__c
            WHERE Status__c = 'Failed'
        ];
        result.syncMeetings = valueOrDefault(settings.Sync_Meetings__c, true);
        result.syncAttendance = valueOrDefault(settings.Sync_Attendance__c, true);
        result.syncRecordings = valueOrDefault(settings.Sync_Recordings__c, true);
        result.syncTranscripts = valueOrDefault(settings.Sync_Transcripts__c, true);
        result.syncWindowDays = settings.Sync_Window_Days__c == null
            ? 90
            : Integer.valueOf(settings.Sync_Window_Days__c);
        return result;
    }

    @AuraEnabled
    public static SetupInfo saveSettings(SetupInfo input) {
        if (input == null) {
            throw new AuraHandledException('Settings are required.');
        }
        if (input.syncWindowDays == null || input.syncWindowDays < 1 || input.syncWindowDays > 365) {
            throw new AuraHandledException('Sync window must be between 1 and 365 days.');
        }
        MS_Teams_Setting__c settings = MS_Teams_Setting__c.getOrgDefaults();
        if (settings.Id == null) {
            settings.SetupOwnerId = UserInfo.getOrganizationId();
        }
        settings.Sync_Meetings__c = input.syncMeetings;
        settings.Sync_Attendance__c = input.syncAttendance;
        settings.Sync_Recordings__c = input.syncRecordings;
        settings.Sync_Transcripts__c = input.syncTranscripts;
        settings.Sync_Window_Days__c = input.syncWindowDays;
        upsert settings;
        return getSetupInfo();
    }

    private static Boolean valueOrDefault(Boolean value, Boolean defaultValue) {
        return value == null ? defaultValue : value;
    }

    public class SetupInfo {
        @AuraEnabled public String tenantName { get; set; }
        @AuraEnabled public String tenantId { get; set; }
        @AuraEnabled public Integer connectedUsers { get; set; }
        @AuraEnabled public Integer activeSubscriptions { get; set; }
        @AuraEnabled public Integer failedSyncs { get; set; }
        @AuraEnabled public Datetime lastSync { get; set; }
        @AuraEnabled public Boolean syncMeetings { get; set; }
        @AuraEnabled public Boolean syncAttendance { get; set; }
        @AuraEnabled public Boolean syncRecordings { get; set; }
        @AuraEnabled public Boolean syncTranscripts { get; set; }
        @AuraEnabled public Integer syncWindowDays { get; set; }
    }
}
