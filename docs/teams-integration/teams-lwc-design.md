# Teams LWC naming and Apex design

Status: proposed design based on the current retrieved Apex and meeting Visualforce page; no new Apex implementation or deployment yet.

## Naming

Use MS Teams as the feature family: `MSTeams` for Apex, `msTeams` for LWC bundles, and `MS_Teams_` for new permission sets and feature configuration. Use `MSGraph` for new shared Microsoft Graph transport/authentication classes only when an actual refactor needs them. Keep the five existing object API names and current field names to preserve data and references. These are naming conventions, not a Salesforce package namespace.

Examples: `MSTeamsMeetingController`, `MSTeamsMeetingService`, `MSTeamsMeetingSelector`, `MSTeamsMeetingRequest`, `msTeamsMeetings`, `msTeamsMeetingForm`, `MS_Teams_User`, `MS_Teams_Admin`. Tests use the production class name plus `Test`.

## Apex responsibilities

| Proposed class | Responsibility / existing source |
| --- | --- |
| MSTeamsMeetingController | Public static AuraEnabled entry points for list, details, create, update, cancel, delete, and explicit synchronization requests. |
| MSTeamsMeetingService | Execute meeting creation and mutation operations against Microsoft Graph and persist meeting state. It is called by the LWC controller, not exposed directly. |
| MSTeamsMeetingSelector | Query schedules through MS_Teams_Record_Link__c, then load participants, intervals, and artifacts with user-mode access enforcement. |
| MSTeamsMeetingController | Resolve any accessible, queryable Salesforce record context and expose the LWC meeting actions. |
| MSTeamsRecordLinkService | Emulate polymorphic WhatId-style association: derive the object type from recordId, validate access, create the unique meeting link, and authorize later meeting actions. |
| MSTeamsGraphClient | Shared named-credential HTTP transport, JSON parsing, path encoding, and Graph error handling. |
| MSTeamsAuthService | Validate delegated and application authentication for a future connection-status UI. |
| MSTeamsAvailabilityService | Find available meeting times for a future scheduling enhancement. |

Create top-level DTO classes as required: MSTeamsMeetingRequest, MSTeamsMeetingSummary, MSTeamsMeetingDetails, MSTeamsParticipant, MSTeamsAttendanceInterval, MSTeamsArtifact, MSTeamsActionResult, MSTeamsConnectionResult, MSTeamsAvailabilityResult. Properties crossing the UI boundary use AuraEnabled. Do not expose the existing nested result classes directly.

## Meeting API contract

- getContext(recordId): supported record context and recipient defaults.
- getMeetings(recordId, pageSize, cursor): paginated stored meetings; cacheable read.
- getMeetingDetails(recordId, scheduleId): stored participants, attendance intervals, artifacts, status, and last sync result; cacheable read.
- createMeeting(request): subject, description, attendees, dates, selected Microsoft time zone, recording/transcription options, and recordId.
- updateMeeting(request): scheduleId and recordId plus supported editable values. Preserve the current backend's limits on description and recording option edits until explicitly implemented.
- cancelMeeting(recordId, scheduleId, reason): cancel externally and retain the local audit record.
- deleteMeeting(recordId, scheduleId, reason): delete the external event while preserving local Deleted status and justification.
- requestSync(recordId, scheduleId): authorize and enqueue the existing sync pipeline, return queued status, then refresh stored data.

Writes and callouts are imperative, non-cacheable actions. Refresh cached lists/details after completion. Scheduling uses an explicit selected time zone rather than the VF page's hard-coded Eastern Standard Time. Prevent duplicate submissions and distinguish external success/local persistence failure to avoid unsafe retries.

## Existing implementation constraints

The LWC boundary is MSTeamsMeetingController and its top-level request/response classes. Service-level inner classes are internal and are not returned to LWC. The implementation has no compile-time dependency on DHS_Lead__c or DHS_Case__c. The existing webhook and lifecycle queueables/services remain in IntelDev because they are background integration endpoints, not LWC controllers; they were removed from this local subset.

User-driven services must check record ownership/context, sharing, CRUD/FLS, and action authorization before Graph side effects. Existing mutation code uses SYSTEM_MODE database updates, so a wrapper alone must not be treated as full access enforcement. Review that path explicitly. Keep privileged webhook/background processing separate from user entry points.

Keep the existing webhook URL and subscription processing during the first migration. Do not rename deployed endpoints or Named Credentials merely to match the naming convention.

## LWC composition

- msTeamsMeetings: record-page container and action coordination.
- msTeamsMeetingForm: create/edit modal with recipients and time-zone-aware scheduling.
- msTeamsMeetingList: list, status, links, and actions.
- msTeamsMeetingDetails: details modal with participant and artifact sections.
- msTeamsParticipants: invitation response and attendance history.
- msTeamsArtifacts: recordings/transcript references and synchronization errors.

The existing five objects and lead-based methods do not establish generic Account/Contact/Opportunity association or chat/content ingestion. Those shared-chat features require separate schema/backend design, not only LWC conversion.

## Validation for implementation

Cover authorized/unauthorized record access, field permissions, create/update/cancel/delete Graph outcomes using HttpCalloutMock, preservation of deleted records, record-context mismatch, empty results, attendance/artifact mapping, explicit time zones, duplicate submissions, and async sync completion. Validate LWC form errors, modal actions, and refresh behavior. No live meeting creation is needed for automated tests.

Sources:
- https://developer.salesforce.com/docs/platform/lwc/guide/apex-expose-method.html
- https://developer.salesforce.com/docs/platform/lwc/guide/apex-security
- https://developer.salesforce.com/docs/platform/lwc/guide/apex-call-imperative
