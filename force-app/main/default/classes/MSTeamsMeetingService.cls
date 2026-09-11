public with sharing class MSTeamsMeetingService {

    private static final String NAMED_CREDENTIAL = 'MicrosoftGraph';
    private static final String GRAPH_API_VERSION = 'v1.0';
    private static final Integer CALLOUT_TIMEOUT = 120000;
    private static final String IMMUTABLE_ID_PREFERENCE =
        'IdType="ImmutableId"';

    /**
     * Creates a calendar meeting where the authenticated Microsoft user
     * is the organizer.
     *
     * Example start/end:
     * startDateTimeValue = '2026-07-24T10:00:00'
     * endDateTimeValue   = '2026-07-24T10:30:00'
     * microsoftTimeZone  = 'Eastern Standard Time'
     */
    public static MeetingResult createMeeting(
        String subject,
        String bodyHtml,
        List<String> attendeeEmails,
        String startDateTimeValue,
        String endDateTimeValue,
        String microsoftTimeZone,
        Boolean createTeamsMeeting
    ) {
        return createMeetingWithContext(
            subject,
            bodyHtml,
            attendeeEmails,
            startDateTimeValue,
            endDateTimeValue,
            microsoftTimeZone,
            createTeamsMeeting,
            false,
            false,
            null
        );
    }

    public static MeetingResult createMeetingForRecord(
        MSTeamsMeetingRequest meetingRequest,
        String bodyHtml,
        String recordLabel
    ) {
        MeetingResult result = createMeetingWithContext(
            meetingRequest.subject,
            bodyHtml,
            meetingRequest.attendeeEmails,
            meetingRequest.startDateTime,
            meetingRequest.endDateTime,
            meetingRequest.microsoftTimeZone,
            true,
            meetingRequest.recordAutomatically == true,
            meetingRequest.enableTranscription == true,
            meetingRequest
        );
        try {
            MSTeamsRecordLinkService.createPrimaryLink(
                meetingRequest.recordId,
                result.scheduleId,
                recordLabel
            );
        } catch (Exception exceptionValue) {
            throw new AuraHandledException(
                'The Microsoft meeting and Salesforce schedule were created, ' +
                'but the source record could not be linked. Microsoft event ID: ' +
                result.eventId + '. ' + exceptionValue.getMessage()
            );
        }
        return result;
    }

    private static void applyRequestedMeetingOptions(
        MeetingResult result,
        Boolean recordAutomatically,
        Boolean enableTranscription,
        MSTeamsMeetingRequest meetingRequest
    ) {
        if (
            result == null ||
            String.isBlank(result.onlineMeetingId)
        ) {
            return;
        }
        Map<String, Object> payload = new Map<String, Object>();
        if (recordAutomatically) {
            payload.put('allowRecording', true);
            payload.put('recordAutomatically', true);
        }
        if (enableTranscription) {
            payload.put('allowTranscription', true);
        }
        if (meetingRequest != null) {
            payload.put(
                'allowMeetingChat',
                meetingRequest.allowMeetingChat == false
                    ? 'disabled'
                    : 'enabled'
            );
            payload.put(
                'allowedPresenters',
                meetingRequest.allowScreenSharing == false
                    ? 'organizer'
                    : 'everyone'
            );
            payload.put(
                'lobbyBypassSettings',
                new Map<String, Object>{
                    'scope' => lobbyBypassScope(meetingRequest)
                }
            );
        }
        if (payload.isEmpty()) {
            return;
        }
        try {
            HttpResponse response = MSTeamsGraphClient.patch(
                '/me/onlineMeetings/' +
                MSTeamsGraphClient.encodePath(
                    result.onlineMeetingId
                ),
                payload
            );
            if (response.getStatusCode() != 200) {
                result.lifecyclePreparationError =
                    buildGraphOperationErrorMessage(
                        response,
                        'apply recording/transcription options'
                    );
            }
        } catch (Exception exceptionValue) {
            result.lifecyclePreparationError =
                'Meeting created, but recording/transcription options ' +
                'could not be applied: ' + exceptionValue.getMessage();
        }
    }

    private static MeetingResult createMeetingWithContext(
        String subject,
        String bodyHtml,
        List<String> attendeeEmails,
        String startDateTimeValue,
        String endDateTimeValue,
        String microsoftTimeZone,
        Boolean createTeamsMeeting,
        Boolean recordAutomatically,
        Boolean enableTranscription,
        MSTeamsMeetingRequest meetingRequest
    ) {
        validateRequest(
            subject,
            attendeeEmails,
            startDateTimeValue,
            endDateTimeValue,
            microsoftTimeZone
        );
        String graphTransactionId = generateTransactionId();

        HttpRequest request = new HttpRequest();
        request.setEndpoint(
            'callout:' + NAMED_CREDENTIAL +
            '/' + GRAPH_API_VERSION +
            '/me/events'
        );
        request.setMethod('POST');
        request.setHeader('Content-Type', 'application/json');
        request.setHeader('Accept', 'application/json');
        request.setHeader('Prefer', IMMUTABLE_ID_PREFERENCE);
        request.setTimeout(CALLOUT_TIMEOUT);
        request.setBody(
            buildRequestBody(
                subject,
                bodyHtml,
                attendeeEmails,
                startDateTimeValue,
                endDateTimeValue,
                microsoftTimeZone,
                createTeamsMeeting == true,
                graphTransactionId,
                meetingRequest
            )
        );

        HttpResponse response;

        try {
            response = new Http().send(request);
        } catch (System.CalloutException calloutException) {
            throw new AuraHandledException(
                'Microsoft Graph callout failed: ' +
                calloutException.getMessage()
            );
        }

        if (response.getStatusCode() != 201) {
            throw new AuraHandledException(
                buildGraphErrorMessage(response)
            );
        }

        MeetingResult result = parseMeetingResult(response.getBody());
        result.transactionId = graphTransactionId;
        if (createTeamsMeeting == true && String.isNotBlank(result.joinUrl)) {
            resolveOnlineMeeting(result);
            applyRequestedMeetingOptions(
                result,
                recordAutomatically == true,
                enableTranscription == true,
                meetingRequest
            );
        }
        result.scheduleId = saveMeetingSchedule(
            subject,
            bodyHtml,
            attendeeEmails,
            startDateTimeValue,
            endDateTimeValue,
            microsoftTimeZone,
            createTeamsMeeting == true,
            result,
            recordAutomatically == true,
            enableTranscription == true,
            meetingRequest
        );

        return result;
    }

    public static void updateScheduledMeeting(
        Id scheduleId,
        Id recordId,
        String subject,
        List<String> attendeeEmails,
        String startDateTimeValue,
        String endDateTimeValue,
        String microsoftTimeZone
    ) {
        validateRequest(
            subject,
            attendeeEmails,
            startDateTimeValue,
            endDateTimeValue,
            microsoftTimeZone
        );
        MS_Team_Call_Schedule__c schedule =
            getActiveSchedule(scheduleId, recordId);

        Map<String, Object> payload = new Map<String, Object>{
            'subject' => subject.trim(),
            'start' => new Map<String, Object>{
                'dateTime' => startDateTimeValue.trim(),
                'timeZone' => microsoftTimeZone.trim()
            },
            'end' => new Map<String, Object>{
                'dateTime' => endDateTimeValue.trim(),
                'timeZone' => microsoftTimeZone.trim()
            },
            'attendees' => buildAttendees(attendeeEmails)
        };
        HttpResponse response = sendEventRequest(
            'PATCH',
            schedule.Graph_Event_ID__c,
            null,
            JSON.serialize(payload)
        );
        if (response.getStatusCode() != 200) {
            throw new AuraHandledException(buildGraphErrorMessage(response));
        }

        schedule.Subject__c = subject.trim();
        schedule.Attendee_Emails__c = String.join(
            normalizeAttendeeEmails(attendeeEmails),
            '; '
        );
        schedule.Start_Date_Time__c = startDateTimeValue.trim();
        schedule.End_Date_Time__c = endDateTimeValue.trim();
        schedule.Microsoft_Time_Zone__c = microsoftTimeZone.trim();
        schedule.Status__c = 'Scheduled';
        schedule.Last_Synchronized_At__c = System.now();
        schedule.Last_Graph_Error__c = null;
        Database.update(schedule, AccessLevel.SYSTEM_MODE);
    }

    public static void cancelScheduledMeeting(
        Id scheduleId,
        Id recordId,
        String cancellationReason
    ) {
        if (String.isBlank(cancellationReason)) {
            throw new AuraHandledException(
                'Cancellation justification is required.'
            );
        }
        MS_Team_Call_Schedule__c schedule =
            getActiveSchedule(scheduleId, recordId);
        HttpResponse response = sendEventRequest(
            'POST',
            schedule.Graph_Event_ID__c,
            'cancel',
            JSON.serialize(new Map<String, Object>{
                'comment' => cancellationReason.trim()
            })
        );
        if (response.getStatusCode() != 202) {
            throw new AuraHandledException(buildGraphErrorMessage(response));
        }

        schedule.Status__c = 'Cancelled';
        schedule.Cancellation_Reason__c = cancellationReason.trim();
        schedule.Cancelled_At__c = System.now();
        schedule.Cancelled_By__c = UserInfo.getUserId();
        schedule.Last_Synchronized_At__c = System.now();
        schedule.Last_Graph_Error__c = null;
        Database.update(schedule, AccessLevel.SYSTEM_MODE);
    }

    public static void deleteScheduledMeeting(
        Id scheduleId,
        Id recordId,
        String deletionJustification
    ) {
        if (String.isBlank(deletionJustification)) {
            throw new AuraHandledException(
                'Deletion justification is required.'
            );
        }
        MS_Team_Call_Schedule__c schedule =
            getActiveSchedule(scheduleId, recordId);
        HttpResponse response = sendEventRequest(
            'DELETE',
            schedule.Graph_Event_ID__c,
            null,
            null
        );
        if (response.getStatusCode() != 204) {
            throw new AuraHandledException(buildGraphErrorMessage(response));
        }

        schedule.Status__c = 'Deleted';
        schedule.Deletion_Justification__c =
            deletionJustification.trim();
        schedule.Deleted_At__c = System.now();
        schedule.Deleted_By__c = UserInfo.getUserId();
        schedule.Last_Synchronized_At__c = System.now();
        schedule.Last_Graph_Error__c = null;
        Database.update(schedule, AccessLevel.SYSTEM_MODE);
    }

    private static MS_Team_Call_Schedule__c getActiveSchedule(
        Id scheduleId,
        Id recordId
    ) {
        if (scheduleId == null || recordId == null) {
            throw new AuraHandledException('Meeting schedule is required.');
        }
        MSTeamsRecordLinkService.assertMeetingLinked(recordId, scheduleId);
        List<MS_Team_Call_Schedule__c> schedules = [
            SELECT
                Id,
                Graph_Event_ID__c,
                Status__c,
                Subject__c,
                Attendee_Emails__c,
                Start_Date_Time__c,
                End_Date_Time__c,
                Microsoft_Time_Zone__c,
                Last_Synchronized_At__c,
                Last_Graph_Error__c,
                Cancellation_Reason__c,
                Cancelled_At__c,
                Cancelled_By__c,
                Deletion_Justification__c,
                Deleted_At__c,
                Deleted_By__c
            FROM MS_Team_Call_Schedule__c
            WHERE Id = :scheduleId
            LIMIT 1
        ];
        if (schedules.isEmpty()) {
            throw new AuraHandledException(
                'The meeting is unavailable for this record.'
            );
        }
        MS_Team_Call_Schedule__c schedule = schedules[0];
        if (String.isBlank(schedule.Graph_Event_ID__c)) {
            throw new AuraHandledException(
                'The Microsoft Graph event ID is missing.'
            );
        }
        if (schedule.Status__c == 'Cancelled' ||
            schedule.Status__c == 'Deleted') {
            throw new AuraHandledException(
                'This meeting is already ' +
                schedule.Status__c.toLowerCase() + '.'
            );
        }
        return schedule;
    }

    private static HttpResponse sendEventRequest(
        String method,
        String eventId,
        String action,
        String body
    ) {
        String encodedEventId =
            EncodingUtil.urlEncode(eventId, 'UTF-8').replace('+', '%20');
        String endpoint = 'callout:' + NAMED_CREDENTIAL +
            '/' + GRAPH_API_VERSION + '/me/events/' + encodedEventId;
        if (String.isNotBlank(action)) {
            endpoint += '/' + action;
        }

        HttpRequest request = new HttpRequest();
        request.setEndpoint(endpoint);
        request.setMethod(method);
        request.setHeader('Accept', 'application/json');
        request.setTimeout(CALLOUT_TIMEOUT);
        if (body != null) {
            request.setHeader('Content-Type', 'application/json');
            request.setBody(body);
        }
        try {
            return new Http().send(request);
        } catch (System.CalloutException calloutException) {
            throw new AuraHandledException(
                'Microsoft Graph callout failed: ' +
                calloutException.getMessage()
            );
        }
    }

    private static String buildRequestBody(
        String subject,
        String bodyHtml,
        List<String> attendeeEmails,
        String startDateTimeValue,
        String endDateTimeValue,
        String microsoftTimeZone,
        Boolean createTeamsMeeting,
        String graphTransactionId,
        MSTeamsMeetingRequest meetingRequest
    ) {
        Map<String, Object> eventPayload = new Map<String, Object>{
            'subject' => subject.trim(),
            'body' => new Map<String, Object>{
                'contentType' => 'HTML',
                'content' => String.isBlank(bodyHtml)
                    ? ''
                    : bodyHtml
            },
            'start' => new Map<String, Object>{
                'dateTime' => startDateTimeValue.trim(),
                'timeZone' => microsoftTimeZone.trim()
            },
            'end' => new Map<String, Object>{
                'dateTime' => endDateTimeValue.trim(),
                'timeZone' => microsoftTimeZone.trim()
            },
            'attendees' => buildAttendees(attendeeEmails),
            'allowNewTimeProposals' => meetingRequest == null ||
                meetingRequest.allowNewTimeProposals != false,
            'responseRequested' => true,
            'transactionId' => graphTransactionId
        };

        if (meetingRequest != null) {
            Integer reminderMinutes = meetingRequest.reminderMinutesBeforeStart;
            if (reminderMinutes != null && reminderMinutes >= 0) {
                eventPayload.put('isReminderOn', true);
                eventPayload.put('reminderMinutesBeforeStart', reminderMinutes);
            } else {
                eventPayload.put('isReminderOn', false);
            }
            if (String.isNotBlank(meetingRequest.location)) {
                eventPayload.put('location', new Map<String, Object>{
                    'displayName' => meetingRequest.location.trim()
                });
            }
        }

        if (createTeamsMeeting) {
            eventPayload.put('isOnlineMeeting', true);
            eventPayload.put(
                'onlineMeetingProvider',
                'teamsForBusiness'
            );
        }

        return JSON.serialize(eventPayload);
    }

    /**
     * Calendar event IDs and Teams online-meeting IDs are different. Artifact
     * APIs (attendance, transcripts and recordings) require the latter.
     * Resolution is deliberately best-effort: the event and Salesforce
     * schedule must still be retained if artifact permissions are not ready.
     */
    private static void resolveOnlineMeeting(MeetingResult result) {
        String encodedJoinUrl = EncodingUtil.urlEncode(
            result.joinUrl,
            'UTF-8'
        ).replace('+', '%20');
        HttpRequest request = new HttpRequest();
        request.setEndpoint(
            'callout:' + NAMED_CREDENTIAL +
            '/' + GRAPH_API_VERSION +
            '/me/onlineMeetings?%24filter=JoinWebUrl%20eq%20%27' +
            encodedJoinUrl +
            '%27'
        );
        request.setMethod('GET');
        request.setHeader('Accept', 'application/json');
        request.setTimeout(CALLOUT_TIMEOUT);

        try {
            HttpResponse response = new Http().send(request);
            if (response.getStatusCode() != 200) {
                result.lifecyclePreparationError =
                    buildGraphOperationErrorMessage(
                        response,
                        'resolve the Teams online meeting'
                    );
                return;
            }

            Map<String, Object> responseData =
                (Map<String, Object>) JSON.deserializeUntyped(
                    response.getBody()
                );
            List<Object> meetings =
                (List<Object>) responseData.get('value');
            if (meetings == null || meetings.isEmpty()) {
                result.lifecyclePreparationError =
                    'The calendar event was created, but Microsoft Graph did ' +
                    'not yet return its Teams online meeting ID.';
                return;
            }

            Map<String, Object> onlineMeeting =
                (Map<String, Object>) meetings[0];
            result.onlineMeetingId =
                getStringValue(onlineMeeting, 'id');
            result.organizerEntraUserId =
                getOrganizerUserId(onlineMeeting);
        } catch (Exception exceptionValue) {
            result.lifecyclePreparationError =
                'The calendar event was created, but the Teams online ' +
                'meeting ID could not be resolved: ' +
                exceptionValue.getMessage();
        }
    }

    private static String getOrganizerUserId(
        Map<String, Object> onlineMeeting
    ) {
        Map<String, Object> participants =
            getMapValue(onlineMeeting, 'participants');
        Map<String, Object> organizer =
            getMapValue(participants, 'organizer');
        Map<String, Object> identity =
            getMapValue(organizer, 'identity');
        Map<String, Object> user =
            getMapValue(identity, 'user');
        return getStringValue(user, 'id');
    }

    private static String generateTransactionId() {
        String value = EncodingUtil.convertToHex(
            Crypto.generateAesKey(128)
        ).toLowerCase();
        return value.substring(0, 8) + '-' +
            value.substring(8, 12) + '-' +
            value.substring(12, 16) + '-' +
            value.substring(16, 20) + '-' +
            value.substring(20);
    }

    private static String lobbyBypassScope(
        MSTeamsMeetingRequest meetingRequest
    ) {
        if (meetingRequest.lobbyRequired == true) {
            return 'organizer';
        }
        if (meetingRequest.meetingAccess == 'Anyone') {
            return 'everyone';
        }
        if (meetingRequest.meetingAccess == 'Organization') {
            return 'organization';
        }
        if (meetingRequest.meetingAccess == 'Organization and Trusted') {
            return 'organizationExcludingGuests';
        }
        return 'invited';
    }

    private static List<Object> buildAttendees(
        List<String> attendeeEmails
    ) {
        List<Object> attendees = new List<Object>();
        for (String normalizedEmail : normalizeAttendeeEmails(attendeeEmails)) {

            attendees.add(
                new Map<String, Object>{
                    'emailAddress' => new Map<String, Object>{
                        'address' => normalizedEmail
                    },
                    'type' => 'required'
                }
            );
        }

        if (attendees.isEmpty()) {
            throw new AuraHandledException(
                'At least one valid attendee email is required.'
            );
        }

        return attendees;
    }

    private static List<String> normalizeAttendeeEmails(
        List<String> attendeeEmails
    ) {
        List<String> normalizedEmails = new List<String>();
        Set<String> processedEmails = new Set<String>();

        for (String attendeeEmail : attendeeEmails) {
            if (String.isBlank(attendeeEmail)) {
                continue;
            }

            String normalizedEmail = attendeeEmail.trim().toLowerCase();

            if (processedEmails.contains(normalizedEmail)) {
                continue;
            }

            validateEmail(normalizedEmail);
            processedEmails.add(normalizedEmail);
            normalizedEmails.add(normalizedEmail);
        }

        return normalizedEmails;
    }

    private static Id saveMeetingSchedule(
        String subject,
        String bodyHtml,
        List<String> attendeeEmails,
        String startDateTimeValue,
        String endDateTimeValue,
        String microsoftTimeZone,
        Boolean createTeamsMeeting,
        MeetingResult result,
        Boolean recordAutomatically,
        Boolean enableTranscription,
        MSTeamsMeetingRequest meetingRequest
    ) {
        MS_Team_Call_Schedule__c schedule =
            new MS_Team_Call_Schedule__c(
                Subject__c = subject.trim(),
                Body_HTML__c = String.isBlank(bodyHtml) ? null : bodyHtml,
                Attendee_Emails__c = String.join(
                    normalizeAttendeeEmails(attendeeEmails),
                    '; '
                ),
                Start_Date_Time__c = startDateTimeValue.trim(),
                End_Date_Time__c = endDateTimeValue.trim(),
                Microsoft_Time_Zone__c = microsoftTimeZone.trim(),
                Teams_Meeting_Requested__c = createTeamsMeeting,
                Recording_Requested__c = recordAutomatically,
                Transcription_Requested__c = enableTranscription,
                Meeting_Type__c = meetingRequest == null
                    ? null
                    : meetingRequest.meetingType,
                Allow_Meeting_Chat__c = meetingRequest == null ||
                    meetingRequest.allowMeetingChat != false,
                Allow_Screen_Sharing__c = meetingRequest == null ||
                    meetingRequest.allowScreenSharing != false,
                Lobby_Required__c = meetingRequest != null &&
                    meetingRequest.lobbyRequired == true,
                Meeting_Access__c = meetingRequest == null
                    ? 'Invited People'
                    : meetingRequest.meetingAccess,
                Teams_Channel_ID__c = meetingRequest == null
                    ? null
                    : meetingRequest.teamsChannelId,
                Location__c = meetingRequest == null
                    ? null
                    : meetingRequest.location,
                Reminder_Minutes_Before_Start__c = meetingRequest == null
                    ? 15
                    : meetingRequest.reminderMinutesBeforeStart,
                Allow_New_Time_Proposals__c = meetingRequest == null ||
                    meetingRequest.allowNewTimeProposals != false,
                Is_Online_Meeting__c = result.isOnlineMeeting,
                Graph_Event_ID__c = result.eventId,
                Graph_Transaction_ID__c = result.transactionId,
                Graph_Change_Key__c = result.changeKey,
                Graph_Last_Modified__c = result.graphLastModified,
                Organizer_Email__c = result.organizerEmail,
                Organizer_Entra_User_ID__c =
                    result.organizerEntraUserId,
                Online_Meeting_ID__c = result.onlineMeetingId,
                Web_Link__c = result.webLink,
                Teams_Join_URL__c = result.joinUrl,
                Status__c = 'Scheduled',
                Graph_Created_At__c = System.now(),
                Last_Synchronized_At__c = System.now(),
                Last_Graph_Error__c = result.lifecyclePreparationError
            );

        try {
            Database.insert(schedule, AccessLevel.SYSTEM_MODE);
        } catch (DmlException dmlException) {
            throw new AuraHandledException(
                'The Microsoft meeting was created, but its Salesforce ' +
                'schedule record could not be saved. Microsoft event ID: ' +
                result.eventId + '. ' + dmlException.getDmlMessage(0)
            );
        }

        return schedule.Id;
    }

    private static void validateRequest(
        String subject,
        List<String> attendeeEmails,
        String startDateTimeValue,
        String endDateTimeValue,
        String microsoftTimeZone
    ) {
        if (String.isBlank(subject)) {
            throw new AuraHandledException(
                'Meeting subject is required.'
            );
        }

        if (subject.trim().length() > 255) {
            throw new AuraHandledException(
                'Meeting subject cannot exceed 255 characters.'
            );
        }

        if (attendeeEmails == null || attendeeEmails.isEmpty()) {
            throw new AuraHandledException(
                'At least one attendee email is required.'
            );
        }

        if (String.isBlank(startDateTimeValue)) {
            throw new AuraHandledException(
                'Meeting start date and time are required.'
            );
        }

        if (String.isBlank(endDateTimeValue)) {
            throw new AuraHandledException(
                'Meeting end date and time are required.'
            );
        }

        if (String.isBlank(microsoftTimeZone)) {
            throw new AuraHandledException(
                'Microsoft time zone is required.'
            );
        }

        validateDateTimeFormat(
            startDateTimeValue,
            'start'
        );

        validateDateTimeFormat(
            endDateTimeValue,
            'end'
        );
    }

    private static void validateDateTimeFormat(
        String dateTimeValue,
        String fieldLabel
    ) {
        String dateTimePattern =
            '^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}(:\\d{2})?$';

        if (!Pattern.matches(
            dateTimePattern,
            dateTimeValue.trim()
        )) {
            throw new AuraHandledException(
                'Invalid ' + fieldLabel +
                ' date/time. Expected format: yyyy-MM-ddTHH:mm:ss.'
            );
        }
    }

    private static void validateEmail(String emailAddress) {
        String emailPattern =
            '^[A-Za-z0-9._%+\\-]+@' +
            '[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}$';

        if (!Pattern.matches(emailPattern, emailAddress)) {
            throw new AuraHandledException(
                'Invalid attendee email: ' + emailAddress
            );
        }
    }

    private static MeetingResult parseMeetingResult(
        String responseBody
    ) {
        Map<String, Object> responseData =
            (Map<String, Object>) JSON.deserializeUntyped(responseBody);

        MeetingResult result = new MeetingResult();
        result.success = true;
        result.message = 'Meeting created successfully.';
        result.eventId = getStringValue(responseData, 'id');
        result.changeKey = getStringValue(responseData, 'changeKey');
        result.graphLastModified = parseGraphDateTime(
            getStringValue(responseData, 'lastModifiedDateTime')
        );
        result.subject = getStringValue(responseData, 'subject');
        result.webLink = getStringValue(responseData, 'webLink');
        result.isOnlineMeeting =
            responseData.get('isOnlineMeeting') == true;

        Map<String, Object> organizer =
            getMapValue(responseData, 'organizer');

        if (organizer != null) {
            Map<String, Object> organizerEmail =
                getMapValue(organizer, 'emailAddress');

            if (organizerEmail != null) {
                result.organizerEmail =
                    getStringValue(organizerEmail, 'address');
            }
        }

        Map<String, Object> onlineMeeting =
            getMapValue(responseData, 'onlineMeeting');

        if (onlineMeeting != null) {
            result.joinUrl =
                getStringValue(onlineMeeting, 'joinUrl');
        }

        Map<String, Object> startValue =
            getMapValue(responseData, 'start');

        if (startValue != null) {
            result.startDateTime =
                getStringValue(startValue, 'dateTime');
            result.timeZone =
                getStringValue(startValue, 'timeZone');
        }

        Map<String, Object> endValue =
            getMapValue(responseData, 'end');

        if (endValue != null) {
            result.endDateTime =
                getStringValue(endValue, 'dateTime');
        }

        return result;
    }

    private static Datetime parseGraphDateTime(String graphDateTimeValue) {
        if (String.isBlank(graphDateTimeValue)) {
            return null;
        }
        try {
            return (Datetime) JSON.deserialize(
                JSON.serialize(graphDateTimeValue),
                Datetime.class
            );
        } catch (Exception ignored) {
            return null;
        }
    }

    private static String buildGraphErrorMessage(
        HttpResponse response
    ) {
        return buildGraphOperationErrorMessage(
            response,
            'create the Microsoft meeting'
        );
    }

    private static String buildGraphOperationErrorMessage(
        HttpResponse response,
        String operationDescription
    ) {
        String graphCode;
        String graphMessage;

        try {
            Map<String, Object> responseData =
                (Map<String, Object>) JSON.deserializeUntyped(
                    response.getBody()
                );

            Map<String, Object> errorData =
                getMapValue(responseData, 'error');

            if (errorData != null) {
                graphCode = getStringValue(errorData, 'code');
                graphMessage = getStringValue(errorData, 'message');
            }
        } catch (Exception parsingException) {
            graphMessage = response.getBody();
        }

        String errorMessage =
            'Unable to ' + operationDescription + '. HTTP ' +
            response.getStatusCode();

        if (!String.isBlank(graphCode)) {
            errorMessage += ' - ' + graphCode;
        }

        if (!String.isBlank(graphMessage)) {
            errorMessage += ': ' + graphMessage;
        }

        return errorMessage;
    }

    private static Map<String, Object> getMapValue(
        Map<String, Object> sourceData,
        String key
    ) {
        if (
            sourceData == null ||
            !sourceData.containsKey(key) ||
            sourceData.get(key) == null
        ) {
            return null;
        }

        return (Map<String, Object>) sourceData.get(key);
    }

    private static String getStringValue(
        Map<String, Object> sourceData,
        String key
    ) {
        if (
            sourceData == null ||
            !sourceData.containsKey(key) ||
            sourceData.get(key) == null
        ) {
            return null;
        }

        return String.valueOf(sourceData.get(key));
    }

    public class MeetingResult {
        public Boolean success { get; set; }
        public String message { get; set; }
        public String eventId { get; set; }
        public String transactionId { get; set; }
        public String changeKey { get; set; }
        public Datetime graphLastModified { get; set; }
        public String subject { get; set; }
        public String organizerEmail { get; set; }
        public String startDateTime { get; set; }
        public String endDateTime { get; set; }
        public String timeZone { get; set; }
        public String webLink { get; set; }
        public Boolean isOnlineMeeting { get; set; }
        public String joinUrl { get; set; }
        public String onlineMeetingId { get; set; }
        public String organizerEntraUserId { get; set; }
        public String lifecyclePreparationError { get; set; }
        public Id scheduleId { get; set; }
    }
}
