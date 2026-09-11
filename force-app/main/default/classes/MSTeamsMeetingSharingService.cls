public without sharing class MSTeamsMeetingSharingService {
    /**
     * Gives active Salesforce users invited by email read access to a meeting.
     * Child participant, attendance, artifact, chat and record-link records
     * inherit this access through their master-detail relationships.
     */
    public static void synchronizeAttendeeShares(
        Id scheduleId,
        String previousAttendeeEmails,
        List<String> currentAttendeeEmails
    ) {
        if (scheduleId == null) {
            return;
        }
        Set<String> previousEmails = normalize(splitEmails(previousAttendeeEmails));
        Set<String> currentEmails = normalize(currentAttendeeEmails);
        Set<String> allEmails = new Set<String>();
        allEmails.addAll(previousEmails);
        allEmails.addAll(currentEmails);
        if (allEmails.isEmpty()) {
            return;
        }

        Map<String, Id> userIdByEmail = new Map<String, Id>();
        for (User userRow : [
            SELECT Id, Email
            FROM User
            WHERE IsActive = true AND Email IN :allEmails
        ]) {
            userIdByEmail.put(userRow.Email.toLowerCase(), userRow.Id);
        }

        Set<Id> previousUserIds = idsForEmails(previousEmails, userIdByEmail);
        Set<Id> currentUserIds = idsForEmails(currentEmails, userIdByEmail);
        Id ownerId = [
            SELECT OwnerId
            FROM MS_Team_Call_Schedule__c
            WHERE Id = :scheduleId
            LIMIT 1
        ].OwnerId;
        previousUserIds.remove(ownerId);
        currentUserIds.remove(ownerId);

        try {
            if (!previousUserIds.isEmpty()) {
                List<MS_Team_Call_Schedule__Share> obsoleteShares = [
                    SELECT Id
                    FROM MS_Team_Call_Schedule__Share
                    WHERE ParentId = :scheduleId
                        AND RowCause = 'Manual'
                        AND UserOrGroupId IN :previousUserIds
                ];
                if (!obsoleteShares.isEmpty()) {
                    Database.delete(obsoleteShares, false);
                }
            }
            if (!currentUserIds.isEmpty()) {
                Set<Id> alreadyShared = new Set<Id>();
                for (MS_Team_Call_Schedule__Share existingShare : [
                    SELECT UserOrGroupId
                    FROM MS_Team_Call_Schedule__Share
                    WHERE ParentId = :scheduleId
                        AND UserOrGroupId IN :currentUserIds
                ]) {
                    alreadyShared.add(existingShare.UserOrGroupId);
                }
                List<MS_Team_Call_Schedule__Share> shares =
                    new List<MS_Team_Call_Schedule__Share>();
                for (Id userId : currentUserIds) {
                    if (!alreadyShared.contains(userId)) {
                        shares.add(new MS_Team_Call_Schedule__Share(
                            ParentId = scheduleId,
                            UserOrGroupId = userId,
                            AccessLevel = 'Read',
                            RowCause = Schema.MS_Team_Call_Schedule__Share.RowCause.Manual
                        ));
                    }
                }
                if (!shares.isEmpty()) {
                    Database.insert(shares, false);
                }
            }
        } catch (Exception ignored) {
            // Meeting creation must remain successful if an attendee cannot be
            // mapped to a shareable internal Salesforce user.
        }
    }

    private static List<String> splitEmails(String value) {
        return String.isBlank(value) ? new List<String>() : value.split('[;,]');
    }

    private static Set<String> normalize(List<String> values) {
        Set<String> result = new Set<String>();
        if (values == null) {
            return result;
        }
        for (String value : values) {
            if (String.isNotBlank(value)) {
                result.add(value.trim().toLowerCase());
            }
        }
        return result;
    }

    private static Set<Id> idsForEmails(
        Set<String> emails,
        Map<String, Id> userIdByEmail
    ) {
        Set<Id> result = new Set<Id>();
        for (String email : emails) {
            if (userIdByEmail.containsKey(email)) {
                result.add(userIdByEmail.get(email));
            }
        }
        return result;
    }
}
