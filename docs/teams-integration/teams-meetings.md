# Teams Meetings Visualforce page

`force-app/main/default/pages/TeamsMeetings.page` contains the meeting CSS, modal markup, JavaScript, and recipient input handler extracted from IntelDev's `DHSLeaddetail` page. The retrieved original remains alongside it for reference.

The page includes Create Meeting and View Meetings buttons, scheduling, recipient chips, recording/transcription options, the meeting list, edit/cancel/delete actions, and the meeting details modal with participant attendance, recordings, and transcripts. Unrelated lead editing, grammar checks, messages, and shared header markup are excluded.

After deployment, open `/apex/TeamsMeetings?id=<DHS_Lead__c record ID>`. The existing controller requires a valid lead ID.

Dependencies already in IntelDev:

- `DHSLeadDetailPageController`, including its lead context and meeting remote actions. It is reused, not copied or modified locally.
- The existing Microsoft Graph integration and MS objects.
- `casedashboard` static resource for the original styles, jQuery, and Bootstrap bundle.
- The original Font Awesome and jquery-confirm CDN assets.

The original scheduling code sends `Eastern Standard Time` to Microsoft Graph; this behavior is preserved. This is an isolated meeting UI, not a self-contained deployment of the backend and its dependencies.

Validation: JavaScript syntax and DOM ID references are checked locally. Salesforce compilation is checked with a page-only dry-run deployment. No meeting is created as part of validation.
