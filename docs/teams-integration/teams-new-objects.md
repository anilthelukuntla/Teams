# MS Teams objects and UI fields

Created locally using the MS_Teams_ object naming convention. Salesforce page/application logic, permission sets, and synchronization services are separate implementation work.

## Screens 1–5 UI fields

The meeting workspace UI adds these fields to `MS_Team_Call_Schedule__c`: `Meeting_Type__c`, `Allow_Meeting_Chat__c`, `Allow_Screen_Sharing__c`, `Lobby_Required__c`, `Meeting_Access__c`, `Teams_Channel_ID__c`, and `Location__c`. Chat and screen sharing default to enabled, lobby defaults to disabled, and meeting access defaults to invited people.

`MS_Team_Meeting_Participant__c` adds `Salesforce_Record_ID__c` and `Salesforce_Object_API_Name__c` so an attendee selected from Salesforce can retain its source identity. These text fields provide polymorphic-style identity without creating lookups to each supported object.

| Object | Custom fields | Relationship and purpose |
| --- | --- | --- |
| MS_Teams_Record_Link__c | 6 | Master-detail to MS_Team_Call_Schedule__c; associates meetings with standard/custom Salesforce records. |
| MS_Teams_Chat_Message__c | 17 | Master-detail to MS_Team_Call_Schedule__c; stores permitted chat content, sender identity, timestamps, deletion status, and provider IDs. |
| MS_Teams_Sync_Log__c | 15 | Optional lookups to meeting and subscription; stores operations, attempts, status, diagnostics, and retry times. Private sharing. |

Record links and chat messages inherit meeting sharing and are deleted if their master meeting is physically deleted. Existing meeting deletion marks status Deleted; that preserves these child records. Sync log lookups clear when a referenced record is physically deleted so logs can remain.

## Record associations

Record_ID__c stores an 18-character Salesforce ID and Object_API_Name__c stores its API name. The LWC passes only recordId. MSTeamsRecordLinkService derives the object API name with `recordId.getSObjectType()`, requires a queryable and accessible object, queries the specific record in user mode, and then creates the link. Reads and meeting mutations repeat the access and link checks. This supports any record type the running user and integration permission set can access without one lookup per object type.

This behaves like Task WhatId from the application's perspective, but Salesforce does not permit a custom polymorphic lookup. Record_ID__c is therefore text rather than a native relationship: Salesforce does not provide automatic referential integrity, related lists on every target object, cascade behavior, or lookup-based reporting. The LWC uses NavigationMixin with Record_ID__c for navigation. Cleanup for deleted target records must be handled by a scheduled process if required.

External_Key__c is a required unique external ID constructed from meeting ID + ':' + target ID (both canonical 18-character IDs). It prevents duplicate links. The first link created with a meeting is Primary; later associations can use Related. The application must enforce at most one Primary role per meeting.

## Chat synchronization

External_Key__c is a required unique external ID: compute the SHA-256 hex digest of a canonical encoded tuple containing meeting ID, Graph chat ID, and Graph message ID. Chat IDs and message IDs are separately required. This supports idempotent upsert across repeated notifications. Preserve provider message type as text and track remote deletion separately from local record deletion.

Message body capacity is 131,072 characters. The future ingestion service must handle over-limit content explicitly rather than silently truncate. HTML content must be sanitized before display. Attachments, reactions, and arbitrary provider payloads are not modeled by this first schema.

## Sync logs

Use one record per attempt with a positive one-based attempt number and a shared correlation ID across retries. Populate start/completion times consistently. Next_Retry_At__c records the planned retry; this field does not itself schedule jobs. Keep diagnostics redacted and configure retention when the sync worker is implemented. No automatic retry or ingestion behavior is added by these objects alone.

## Validation and deployment

Salesforce metadata validation succeeded against IntelDev using a dry-run deployment of manifest/teams-new-objects.xml. No changes were deployed to IntelDev. The two existing package manifests also include the new objects.
