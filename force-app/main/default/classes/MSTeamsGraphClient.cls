public with sharing class MSTeamsGraphClient {
    private static final String NAMED_CREDENTIAL = 'MicrosoftGraph';
    private static final String API_VERSION = 'v1.0';
    private static final Integer TIMEOUT = 120000;

    public static HttpResponse get(String relativePath) {
        return send('GET', relativePath, null);
    }

    public static HttpResponse post(String relativePath, Object body) {
        return send('POST', relativePath, body);
    }

    public static HttpResponse patch(String relativePath, Object body) {
        return send('PATCH', relativePath, body);
    }

    public static HttpResponse send(
        String method,
        String relativePath,
        Object body
    ) {
        String normalizedPath = relativePath.startsWith('/')
            ? relativePath
            : '/' + relativePath;
        HttpRequest request = new HttpRequest();
        request.setEndpoint(
            'callout:' + NAMED_CREDENTIAL +
            '/' + API_VERSION + normalizedPath
        );
        request.setMethod(method);
        request.setHeader('Accept', 'application/json');
        request.setHeader('Prefer', 'IdType="ImmutableId"');
        request.setTimeout(TIMEOUT);
        if (body != null) {
            request.setHeader('Content-Type', 'application/json');
            request.setBody(
                body instanceof String
                    ? (String) body
                    : JSON.serialize(body)
            );
        }
        return new Http().send(request);
    }

    public static Map<String, Object> requireJson(
        HttpResponse response,
        Set<Integer> successCodes,
        String operation
    ) {
        if (!successCodes.contains(response.getStatusCode())) {
            throw new GraphException(
                operation + ' failed. HTTP ' +
                response.getStatusCode() + ': ' +
                response.getBody()
            );
        }
        if (String.isBlank(response.getBody())) {
            return new Map<String, Object>();
        }
        return (Map<String, Object>) JSON.deserializeUntyped(
            response.getBody()
        );
    }

    public static String encodePath(String value) {
        return EncodingUtil.urlEncode(value, 'UTF-8').replace('+', '%20');
    }

    public static Datetime parseDateTime(Object value) {
        if (value == null || String.isBlank(String.valueOf(value))) {
            return null;
        }
        try {
            return (Datetime) JSON.deserialize(
                JSON.serialize(String.valueOf(value)),
                Datetime.class
            );
        } catch (Exception ignored) {
            return null;
        }
    }

    public class GraphException extends Exception {
    }
}