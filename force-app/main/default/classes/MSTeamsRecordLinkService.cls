public with sharing class MSTeamsRecordLinkService {
    public static MSTeamsContextResult getContext(Id recordId) {
        Schema.DescribeSObjectResult describe = validateRecordType(recordId);
        String objectApiName = describe.getName();
        String nameField = findNameField(describe);
        String emailField = findEmailField(describe);
        List<String> selectFields = new List<String>{ 'Id' };
        if (nameField != null) {
            selectFields.add(nameField);
        }
        if (emailField != null && emailField != nameField) {
            selectFields.add(emailField);
        }
        List<SObject> rows = Database.queryWithBinds(
            'SELECT ' + String.join(selectFields, ',') +
            ' FROM ' + objectApiName + ' WHERE Id = :recordId LIMIT 1',
            new Map<String, Object>{ 'recordId' => recordId },
            AccessLevel.USER_MODE
        );
        if (rows.isEmpty()) {
            throw new AuraHandledException('Record is unavailable.');
        }
        SObject row = rows[0];
        String recordName = nameField == null
            ? String.valueOf(recordId)
            : String.valueOf(row.get(nameField));
        MSTeamsContextResult result = new MSTeamsContextResult();
        result.recordId = recordId;
        result.objectApiName = objectApiName;
        result.recordName = recordName;
        result.defaultRecipientEmails = new List<String>();
        if (emailField != null && row.get(emailField) != null) {
            result.defaultRecipientEmails.add(
                String.valueOf(row.get(emailField)).trim().toLowerCase()
            );
        }
        result.defaultSubject = describe.getLabel() + ' meeting - ' + recordName;
        result.defaultDescription = 'Microsoft Teams meeting for ' +
            describe.getLabel() + ' ' + recordName + '.';
        return result;
    }

    public static Set<Id> getMeetingIds(Id recordId) {
        validateRecordAccess(recordId);
        Set<Id> meetingIds = new Set<Id>();
        String canonicalId = canonicalId(recordId);
        for (MS_Teams_Record_Link__c link : [
            SELECT Meeting_Schedule__c
            FROM MS_Teams_Record_Link__c
            WHERE Record_ID__c = :canonicalId
            WITH USER_MODE
        ]) {
            meetingIds.add(link.Meeting_Schedule__c);
        }
        return meetingIds;
    }

    public static void assertMeetingLinked(Id recordId, Id scheduleId) {
        validateRecordAccess(recordId);
        String canonicalId = canonicalId(recordId);
        Integer links = [
            SELECT COUNT()
            FROM MS_Teams_Record_Link__c
            WHERE Record_ID__c = :canonicalId
                AND Meeting_Schedule__c = :scheduleId
            WITH USER_MODE
        ];
        if (links != 1) {
            throw new AuraHandledException(
                'Meeting is unavailable for this record.'
            );
        }
    }

    public static void createPrimaryLink(
        Id recordId,
        Id scheduleId,
        String recordLabel
    ) {
        Schema.DescribeSObjectResult describe = validateRecordType(recordId);
        MS_Teams_Record_Link__c link = new MS_Teams_Record_Link__c(
            Meeting_Schedule__c = scheduleId,
            Record_ID__c = canonicalId(recordId),
            Object_API_Name__c = describe.getName(),
            Record_Label__c = recordLabel,
            Relationship_Role__c = 'Primary',
            External_Key__c = canonicalId(scheduleId) + ':' +
                canonicalId(recordId)
        );
        Database.insert(link, AccessLevel.USER_MODE);
    }

    public static void validateRecordAccess(Id recordId) {
        getContext(recordId);
    }

    private static Schema.DescribeSObjectResult validateRecordType(Id recordId) {
        if (recordId == null) {
            throw new AuraHandledException('recordId is required.');
        }
        Schema.DescribeSObjectResult describe =
            recordId.getSObjectType().getDescribe();
        if (!describe.isQueryable() || !describe.isAccessible()) {
            throw new AuraHandledException(
                'This record type is not available for Teams meetings.'
            );
        }
        return describe;
    }

    private static String findNameField(
        Schema.DescribeSObjectResult describe
    ) {
        for (Schema.SObjectField fieldToken :
            describe.fields.getMap().values()) {
            Schema.DescribeFieldResult fieldDescribe = fieldToken.getDescribe();
            if (fieldDescribe.isNameField() && fieldDescribe.isAccessible()) {
                return fieldDescribe.getName();
            }
        }
        return null;
    }

    private static String findEmailField(
        Schema.DescribeSObjectResult describe
    ) {
        Map<String, Schema.SObjectField> fields = describe.fields.getMap();
        for (String candidate : new List<String>{ 'Email', 'PersonEmail' }) {
            if (
                fields.containsKey(candidate) &&
                fields.get(candidate).getDescribe().isAccessible()
            ) {
                return candidate;
            }
        }
        return null;
    }

    private static String canonicalId(Id recordId) {
        return String.valueOf(recordId);
    }
}
