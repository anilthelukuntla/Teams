@IsTest
private class MSTeamsOptionsTest {
    private class OptionsMock implements HttpCalloutMock {
        public Map<String, Object> sent;
        public HttpResponse respond(HttpRequest request) {
            System.assertEquals('PATCH', request.getMethod());
            sent = (Map<String, Object>) JSON.deserializeUntyped(request.getBody());
            HttpResponse response = new HttpResponse();
            response.setStatusCode(200);
            response.setBody('{}');
            return response;
        }
    }

    @IsTest
    static void explicitOffReachesMicrosoft() {
        MSTeamsMeetingRequest request = new MSTeamsMeetingRequest();
        request.allowRecording = false;
        request.allowTranscription = false;
        OptionsMock mock = applyOptions(request, false, false);
        System.assertEquals(false, mock.sent.get('allowRecording'));
        System.assertEquals(false, mock.sent.get('recordAutomatically'));
        System.assertEquals(false, mock.sent.get('allowTranscription'));
    }

    @IsTest
    static void automaticFeaturesRequirePermission() {
        MSTeamsMeetingRequest request = new MSTeamsMeetingRequest();
        request.allowRecording = false;
        request.allowTranscription = false;
        OptionsMock mock = applyOptions(request, true, true);
        System.assertEquals(true, mock.sent.get('allowRecording'));
        System.assertEquals(true, mock.sent.get('recordAutomatically'));
        System.assertEquals(true, mock.sent.get('allowTranscription'));
    }

    @IsTest
    static void omittedPreferencesPreserveLegacyBehavior() {
        OptionsMock mock = applyOptions(new MSTeamsMeetingRequest(), false, false);
        System.assertEquals(false, mock.sent.containsKey('allowRecording'));
        System.assertEquals(false, mock.sent.containsKey('allowTranscription'));
    }

    private static OptionsMock applyOptions(MSTeamsMeetingRequest request, Boolean recordAutomatically, Boolean transcript) {
        OptionsMock mock = new OptionsMock();
        Test.setMock(HttpCalloutMock.class, mock);
        MSTeamsMeetingService.MeetingResult result = new MSTeamsMeetingService.MeetingResult();
        result.onlineMeetingId = 'test-meeting';
        MSTeamsMeetingService.applyRequestedMeetingOptions(result, recordAutomatically, transcript, request);
        System.assertEquals(null, result.lifecyclePreparationError);
        System.assertNotEquals(null, mock.sent);
        return mock;
    }
}
