@IsTest
private class MSTeamsAuthServiceTest {
    @IsTest
    static void returnsConnectedMicrosoftUser() {
        Test.setMock(
            HttpCalloutMock.class,
            new GraphProfileMock(
                200,
                '{"id":"entra-123","displayName":"James Taylor",' +
                '"userPrincipalName":"james.taylor@company.com",' +
                '"mail":"james.taylor@company.com"}'
            )
        );

        Test.startTest();
        MSTeamsConnectionResult result =
            MSTeamsMeetingController.getMicrosoftConnectionStatus();
        Test.stopTest();

        System.assertEquals(true, result.connected);
        System.assertEquals('Connected', result.status);
        System.assertEquals(200, result.statusCode);
        System.assertEquals('entra-123', result.microsoftUserId);
        System.assertEquals('James Taylor', result.displayName);
        System.assertEquals('james.taylor@company.com', result.email);
    }

    @IsTest
    static void returnsDisconnectedForUnauthorizedToken() {
        Test.setMock(
            HttpCalloutMock.class,
            new GraphProfileMock(
                401,
                '{"error":{"code":"InvalidAuthenticationToken",' +
                '"message":"Access token is missing or invalid."}}'
            )
        );

        Test.startTest();
        MSTeamsConnectionResult result =
            MSTeamsMeetingController.getMicrosoftConnectionStatus();
        Test.stopTest();

        System.assertEquals(false, result.connected);
        System.assertEquals('Disconnected', result.status);
        System.assertEquals(401, result.statusCode);
        System.assert(
            result.message.contains('Microsoft Graph authentication failed')
        );
    }

    private class GraphProfileMock implements HttpCalloutMock {
        private Integer statusCode;
        private String body;

        GraphProfileMock(Integer statusCode, String body) {
            this.statusCode = statusCode;
            this.body = body;
        }

        public HttpResponse respond(HttpRequest request) {
            System.assertEquals('GET', request.getMethod());
            System.assert(
                request.getEndpoint().endsWith(
                    '/v1.0/me?$select=id,displayName,userPrincipalName,mail'
                )
            );
            HttpResponse response = new HttpResponse();
            response.setStatusCode(statusCode);
            response.setStatus(
                statusCode == 200 ? 'OK' : 'Unauthorized'
            );
            response.setHeader('Content-Type', 'application/json');
            response.setBody(body);
            return response;
        }
    }
}
