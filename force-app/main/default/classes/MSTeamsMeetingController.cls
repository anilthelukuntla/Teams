public with sharing class MSTeamsMeetingController {
    @AuraEnabled
    public static MSTeamsConnectionResult getMicrosoftConnectionStatus() {
        return toConnectionResult(
            MSTeamsAuthService.validateDelegatedAuthentication()
        );
    }

    @AuraEnabled
    public static String getMicrosoftAuthorizationUrl() {
        return MSTeamsAuthService.getDelegatedAuthorizationUrl();
    }

    @AuraEnabled(cacheable=true)
    public static MSTeamsContextResult getContext(Id recordId) {
        return MSTeamsRecordLinkService.getContext(recordId);
    }

    @AuraEnabled(cacheable=true)
    public static List<MSTeamsMeetingSummary> getMeetings(Id recordId) {
        MSTeamsRecordLinkService.validateRecordAccess(recordId);
        return MSTeamsMeetingSelector.getMeetings(recordId);
    }

    @AuraEnabled(cacheable=true)
    public static List<MSTeamsAttendeeOption> searchAttendees(
        Id recordId,
        String searchTerm
    ) {
        MSTeamsRecordLinkService.validateRecordAccess(recordId);
        return MSTeamsAttendeeSearchService.search(searchTerm);
    }

    @AuraEnabled(cacheable=true)
    public static MSTeamsMeetingDetails getMeetingDetails(
        Id recordId,
        Id scheduleId
    ) {
        MSTeamsRecordLinkService.validateRecordAccess(recordId);
        requireId(scheduleId, 'scheduleId');
        return MSTeamsMeetingSelector.getMeetingDetails(recordId, scheduleId);
    }

    @AuraEnabled
    public static String getTranscriptContent(Id recordId, Id scheduleId) {
        MSTeamsRecordLinkService.assertMeetingLinked(recordId, scheduleId);
        requireMicrosoftConnection();
        List<MS_Team_Meeting_Artifact__c> artifacts = [
            SELECT Content_URL__c, Processing_Status__c
            FROM MS_Team_Meeting_Artifact__c
            WHERE Meeting_Schedule__c = :scheduleId
                AND Artifact_Type__c = 'Transcript'
            WITH USER_MODE
            ORDER BY Created_At__c DESC
            LIMIT 1
        ];
        if (artifacts.isEmpty() || String.isBlank(artifacts[0].Content_URL__c)) {
            throw new AuraHandledException('The transcript is not available yet.');
        }
        String contentPath = artifacts[0].Content_URL__c.trim();
        contentPath = contentPath.replaceFirst('^https://graph.microsoft.com/v1\\.0', '');
        contentPath = contentPath.replaceFirst('^/v1\\.0', '');
        HttpResponse response;
        try {
            response = MSTeamsGraphClient.get(contentPath);
        } catch (Exception exceptionValue) {
            throw new AuraHandledException('Microsoft Graph could not retrieve the transcript: ' + exceptionValue.getMessage());
        }
        if (response.getStatusCode() != 200) {
            throw new AuraHandledException('Microsoft Graph could not retrieve the transcript. HTTP ' + response.getStatusCode() + '.');
        }
        return response.getBody();
    }

    @AuraEnabled
    public static MSTeamsActionResult createMeeting(
        MSTeamsMeetingRequest request
    ) {
        requireMeetingCreateAccess();
        validateMeetingRequest(request, false);
        requireMicrosoftConnection();
        MSTeamsMeetingService.MeetingResult serviceResult =
            MSTeamsMeetingService.createMeetingForRecord(
                request,
                descriptionToHtml(request.description),
                MSTeamsRecordLinkService.getContext(request.recordId).recordName
            );
        MSTeamsActionResult result = MSTeamsActionResult.succeeded(
            serviceResult.scheduleId,
            serviceResult.message
        );
        result.eventId = serviceResult.eventId;
        result.joinUrl = serviceResult.joinUrl;
        result.webLink = serviceResult.webLink;
        return result;
    }

    @AuraEnabled
    public static MSTeamsActionResult updateMeeting(
        MSTeamsMeetingRequest request
    ) {
        requireMeetingEditAccess();
        validateMeetingRequest(request, true);
        validateScheduleAccess(request.recordId, request.scheduleId);
        requireMicrosoftConnection();
        MSTeamsMeetingService.updateScheduledMeeting(
            request.scheduleId,
            request.recordId,
            request.subject,
            request.attendeeEmails,
            request.startDateTime,
            request.endDateTime,
            request.microsoftTimeZone
        );
        return MSTeamsActionResult.succeeded(
            request.scheduleId,
            'Teams meeting updated successfully.'
        );
    }

    @AuraEnabled
    public static MSTeamsActionResult cancelMeeting(
        MSTeamsActionRequest request
    ) {
        requireMeetingEditAccess();
        validateActionRequest(request);
        requireMicrosoftConnection();
        MSTeamsMeetingService.cancelScheduledMeeting(
            request.scheduleId,
            request.recordId,
            request.reason
        );
        return MSTeamsActionResult.succeeded(
            request.scheduleId,
            'Teams meeting cancelled successfully.'
        );
    }

    @AuraEnabled
    public static MSTeamsActionResult deleteMeeting(
        MSTeamsActionRequest request
    ) {
        requireMeetingEditAccess();
        validateActionRequest(request);
        requireMicrosoftConnection();
        MSTeamsMeetingService.deleteScheduledMeeting(
            request.scheduleId,
            request.recordId,
            request.reason
        );
        return MSTeamsActionResult.succeeded(
            request.scheduleId,
            'Teams meeting deleted successfully.'
        );
    }

    private static void validateMeetingRequest(
        MSTeamsMeetingRequest request,
        Boolean scheduleRequired
    ) {
        if (request == null) {
            throw new AuraHandledException('Meeting request is required.');
        }
        MSTeamsRecordLinkService.validateRecordAccess(request.recordId);
        if (scheduleRequired) {
            requireId(request.scheduleId, 'scheduleId');
        }
        if (String.isBlank(request.subject)) {
            throw new AuraHandledException('Meeting subject is required.');
        }
        if (request.attendeeEmails == null || request.attendeeEmails.isEmpty()) {
            throw new AuraHandledException('At least one attendee is required.');
        }
        if (
            String.isBlank(request.startDateTime) ||
            String.isBlank(request.endDateTime) ||
            String.isBlank(request.microsoftTimeZone)
        ) {
            throw new AuraHandledException(
                'Start, end, and Microsoft time zone are required.'
            );
        }
    }

    private static void validateActionRequest(MSTeamsActionRequest request) {
        if (request == null) {
            throw new AuraHandledException('Action request is required.');
        }
        MSTeamsRecordLinkService.validateRecordAccess(request.recordId);
        requireId(request.scheduleId, 'scheduleId');
        validateScheduleAccess(request.recordId, request.scheduleId);
        if (String.isBlank(request.reason)) {
            throw new AuraHandledException('A reason is required.');
        }
    }

    private static void validateScheduleAccess(Id recordId, Id scheduleId) {
        MSTeamsRecordLinkService.assertMeetingLinked(recordId, scheduleId);
    }

    private static MSTeamsAuthService.AuthenticationResult
        requireMicrosoftConnection() {
        MSTeamsAuthService.AuthenticationResult authResult =
            MSTeamsAuthService.validateDelegatedAuthentication();
        if (authResult.authenticated != true) {
            String detail = String.isBlank(authResult.message)
                ? 'Connect your Microsoft 365 account before continuing.'
                : authResult.message;
            throw new AuraHandledException(
                'Microsoft 365 connection required. ' + detail
            );
        }
        return authResult;
    }

    private static MSTeamsConnectionResult toConnectionResult(
        MSTeamsAuthService.AuthenticationResult authResult
    ) {
        MSTeamsConnectionResult result = new MSTeamsConnectionResult();
        result.connected = authResult.authenticated == true;
        result.statusCode = authResult.statusCode;
        result.status = result.connected ? 'Connected' : 'Disconnected';
        result.message = authResult.message;
        result.microsoftUserId = authResult.microsoftUserId;
        result.displayName = authResult.displayName;
        result.userPrincipalName = authResult.userPrincipalName;
        result.email = String.isBlank(authResult.email)
            ? authResult.userPrincipalName
            : authResult.email;
        return result;
    }

    private static void requireId(Id value, String parameterName) {
        if (value == null) {
            throw new AuraHandledException(parameterName + ' is required.');
        }
    }

    private static void requireMeetingCreateAccess() {
        if (!MS_Team_Call_Schedule__c.SObjectType.getDescribe().isCreateable()) {
            throw new AuraHandledException(
                'You do not have permission to schedule Microsoft Teams meetings.'
            );
        }
    }

    private static void requireMeetingEditAccess() {
        if (!MS_Team_Call_Schedule__c.SObjectType.getDescribe().isUpdateable()) {
            throw new AuraHandledException(
                'You do not have permission to manage Microsoft Teams meetings.'
            );
        }
    }

    private static String descriptionToHtml(String description) {
        if (String.isBlank(description)) {
            return '';
        }
        return '<p>' + description.escapeHtml4()
            .replace('\r\n', '<br/>')
            .replace('\n', '<br/>') + '</p>';
    }
}
