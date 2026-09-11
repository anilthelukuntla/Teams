# Microsoft 365 per-user authentication flow

```mermaid
sequenceDiagram
    participant U as Salesforce User
    participant L as Teams LWC
    participant A as Apex
    participant S as Salesforce Per-User External Credential
    participant M as Microsoft Entra / Graph

    L->>A: Check Microsoft connection
    A->>M: GET /v1.0/me
    M-->>A: Missing user credential
    A-->>L: Disconnected
    L-->>U: Show Connect Microsoft 365
    U->>L: Click Connect Microsoft 365
    L->>A: Request authorization URL
    A->>S: Generate per-user OAuth URL
    S-->>A: Authorization URL
    A-->>L: Return URL
    L->>M: Open Microsoft sign-in
    U->>M: Sign in and grant consent
    M->>S: OAuth callback
    S->>S: Store encrypted user tokens
    L->>A: Recheck connection
    A->>M: GET /v1.0/me with user token
    M-->>A: Microsoft user identity
    A-->>L: Connected name and email
    L-->>U: Enable Schedule Meeting
```

Interactive meeting operations use `callout:MicrosoftGraph/v1.0/me/...`. Salesforce attaches the current user's encrypted Microsoft OAuth token, so Microsoft creates the event with that user as its organizer.
