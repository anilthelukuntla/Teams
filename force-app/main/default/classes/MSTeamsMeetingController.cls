public with sharing class MSTeamsMeetingController {
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
    public static MSTeamsMeetingDetails getMeetingDetails(
        Id recordId,
        Id scheduleId
    ) {
        MSTeamsRecordLinkService.validateRecordAccess(recordId);
        requireId(scheduleId, 'scheduleId');
        return MSTeamsMeetingSelector.getMeetingDetails(recordId, scheduleId);
    }

    @AuraEnabled
    public static MSTeamsActionResult createMeeting(
        MSTeamsMeetingRequest request
    ) {
        validateMeetingRequest(request, false);
        MSTeamsMeetingService.MeetingResult serviceResult =
            MSTeamsMeetingService.createMeetingForRecord(
                request.subject,
                descriptionToHtml(request.description),
                request.attendeeEmails,
                request.startDateTime,
                request.endDateTime,
                request.microsoftTimeZone,
                request.recordAutomatically == true,
                request.enableTranscription == true,
                request.recordId,
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
        validateMeetingRequest(request, true);
        validateScheduleAccess(request.recordId, request.scheduleId);
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
        validateActionRequest(request);
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
        validateActionRequest(request);
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

    private static void requireId(Id value, String parameterName) {
        if (value == null) {
            throw new AuraHandledException(parameterName + ' is required.');
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
