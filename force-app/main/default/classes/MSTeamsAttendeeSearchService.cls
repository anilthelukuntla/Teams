public with sharing class MSTeamsAttendeeSearchService {
    public static List<MSTeamsAttendeeOption> search(String searchTerm) {
        String normalized = String.isBlank(searchTerm) ? '' : searchTerm.trim();
        if (normalized.length() < 2) {
            return new List<MSTeamsAttendeeOption>();
        }
        String likeTerm = '%' + String.escapeSingleQuotes(normalized) + '%';
        List<MSTeamsAttendeeOption> results = new List<MSTeamsAttendeeOption>();
        for (Contact row : [
            SELECT Id, Name, Email, Title
            FROM Contact
            WHERE Email != null AND (Name LIKE :likeTerm OR Email LIKE :likeTerm)
            WITH USER_MODE ORDER BY Name LIMIT 15
        ]) {
            results.add(toOption(row.Id, 'Contact', row.Name, row.Title, row.Email));
        }
        for (Lead row : [
            SELECT Id, Name, Email, Title
            FROM Lead
            WHERE Email != null AND IsConverted = false
                AND (Name LIKE :likeTerm OR Email LIKE :likeTerm)
            WITH USER_MODE ORDER BY Name LIMIT 10
        ]) {
            results.add(toOption(row.Id, 'Lead', row.Name, row.Title, row.Email));
        }
        for (User row : [
            SELECT Id, Name, Email, Title
            FROM User
            WHERE Email != null AND IsActive = true
                AND (Name LIKE :likeTerm OR Email LIKE :likeTerm)
            WITH USER_MODE ORDER BY Name LIMIT 10
        ]) {
            results.add(toOption(row.Id, 'User', row.Name, row.Title, row.Email));
        }
        return results;
    }

    private static MSTeamsAttendeeOption toOption(
        Id recordId,
        String objectApiName,
        String displayName,
        String title,
        String email
    ) {
        MSTeamsAttendeeOption option = new MSTeamsAttendeeOption();
        option.recordId = recordId;
        option.objectApiName = objectApiName;
        option.displayName = displayName;
        option.title = title;
        option.email = email;
        option.sourceLabel = objectApiName;
        return option;
    }
}
