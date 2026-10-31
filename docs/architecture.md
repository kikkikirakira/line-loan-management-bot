# Architecture

```mermaid
flowchart LR
    U[LINE user] --> L[LINE Messaging API]
    L --> P[Apps Script doPost]
    P --> S[(Google Sheets)]
    T[Time-driven trigger] --> R[Overdue reminder check]
    R --> S
    R --> L
    B[Web browser] --> G[Apps Script doGet / Dashboard]
    G --> S
    C[Script Properties] --> P
    C --> G
    C --> R
```

The application separates configuration from source code. `SPREADSHEET_ID`, `CHANNEL_ACCESS_TOKEN`, and other deployment-specific values are stored in Apps Script Script Properties rather than committed to Git.
