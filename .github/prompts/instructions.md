You are a senior Microsoft 365 / Azure engineer. Generate a complete, working demo application (repo) that implements:

GOAL
A Microsoft Word Task Pane Office Add-in with a Blazor WebAssembly UI that:
1) Authenticates the signed-in Office user using Office SSO (Office.auth.getAccessToken / OfficeRuntime.auth.getAccessToken).
2) Uses the signed-in user identity (UPN / preferred_username from the token claims) to query a SQL database for snippets owned/allowed for that user.
3) Lets the user pick either:
   - Text snippets (insert into Word at cursor)
   - Image snippets (insert into Word at cursor as inline pictures from Base64)
4) Includes secure backend API + Azure infrastructure scripts + thorough admin documentation.

IMPORTANT ARCHITECTURE / SECURITY REQUIREMENTS
- The add-in runs as an Office Web Add-in (task pane) and must NOT connect directly to SQL from the browser/task pane.
- Use Office SSO token acquisition on the client to identify the user. Office.auth.getAccessToken is for obtaining a token for the add-in’s web app, and server-side code can validate it and (if needed) use OAuth On-Behalf-Of flow for downstream APIs. Implement the safe pattern: client gets token -> calls backend with Bearer token -> backend validates token and authorizes user. (Reference: Office.auth.getAccessToken behavior and OBO concepts.) 
- For database connectivity, default to Azure SQL + Managed Identity for the backend (no secrets). Document that Entra ID / Managed Identity is not supported for on-prem SQL Server (so the demo targets Azure SQL). 
- Insert images into Word using Word JavaScript API insertInlinePictureFromBase64.
- Provide both (A) Infrastructure as Code and (B) step-by-step manual admin instructions for Azure resources and Entra app registrations.
- Include deployment options for Office add-ins: Centralized Deployment and SharePoint app catalog/trusted catalog methods (including mention that sovereign/government clouds may require alternative guidance).

TECH STACK
Repo uses a monorepo layout:
- /add-in/      (Office add-in: manifest, task pane web assets, Word JS interactions)
- /blazor-ui/   (Blazor WebAssembly UI that runs inside the task pane)
- /api/         (ASP.NET Core Web API for data access + auth)
- /infra/       (Bicep scripts + optional Azure CLI scripts)
- /docs/        (Admin + developer docs, diagrams, runbooks)

OFFICE ADD-IN REQUIREMENTS
1) Add-in type: Word Task Pane add-in.
2) Manifest:
   - Include proper HTTPS URLs for task pane and commands.
   - Include required permissions/scopes for SSO configuration.
3) Client auth:
   - Use Office.auth.getAccessToken (or OfficeRuntime.auth.getAccessToken) with allowSignInPrompt/allowConsentPrompt.
   - Send the token in Authorization: Bearer header to the backend API.
   - Include robust error handling.
4) Word integration:
   - Insert selected text snippet at current selection.
   - Insert selected image snippet at current selection as inline picture from Base64.
   - Use Word.run batching pattern and ensure context.sync calls.
5) UI:
   - Two tabs: “Text Snippets” and “Image Snippets”.
   - Search box + list/grid.
   - “Insert” button on each item.
   - Show signed-in user and a “Refresh” button.

BACKEND API REQUIREMENTS (ASP.NET Core)
1) Authentication/Authorization:
   - Validate the incoming access token from Office SSO (JWT validation).
   - Extract user identifier from claims (prefer preferred_username or upn) and use it for filtering.
   - Implement basic authorization rules: only return snippets where OwnerUpn matches caller OR where caller is in an AllowedUpns table.
2) Endpoints:
   - GET /api/me  -> returns normalized user info (upn, name, oid, tid)
   - GET /api/text-snippets -> returns snippets for caller
   - GET /api/image-snippets -> returns snippets for caller
   - (Optional) GET /api/image-snippets/{id}/content -> returns base64
3) Data access:
   - Use Azure SQL.
   - Use Managed Identity from the API host to authenticate to Azure SQL.
   - Use minimal privileges, read-only for demo.
   - Use EF Core with migrations or SQL scripts (include both).
4) Data model (minimum):
   - TextSnippet(Id, OwnerUpn, Title, BodyText, Tags, UpdatedUtc)
   - ImageSnippet(Id, OwnerUpn, Title, ImageBase64, MimeType, Tags, UpdatedUtc)
   - AllowedUser(OwnerUpn, AllowedUpn)  (optional sharing)
5) Provide seed script with sample data for 2-3 users.

INFRASTRUCTURE REQUIREMENTS
Provide /infra with:
A) Bicep templates that deploy:
- Resource group (optional)
- App Service (or Azure Container Apps) for the API, with system-assigned managed identity enabled
- Azure SQL Server + Azure SQL Database
- Configure Entra admin for SQL (document manual step if needed)
- Key Vault optional (only if you must store something; prefer no secrets)
- CORS settings and allowed origins for the add-in host

B) Azure CLI scripts that:
- Deploy the Bicep
- Enable managed identity on the API host
- Grant the managed identity access to Azure SQL (create contained user from external provider; grant db_datareader; optional db_datawriter for demo)
- Output necessary URLs

APP REGISTRATIONS / ENTRA ID REQUIREMENTS
Provide /docs with BOTH:
1) Manual admin steps for Entra app registration(s), including:
   - SPA registration for the add-in front-end (redirect URIs for add-in task pane, dialog if used)
   - API app registration for the backend (Expose an API, define scopes, add application ID URI)
   - Configure required permissions and admin consent
   - Configure manifest webApplicationInfo fields in the Office add-in manifest accordingly
2) Optional automation script (PowerShell or Azure CLI where possible) to create these registrations and configure them.
3) Clear explanation of how Office SSO token is obtained on client and validated on server.

DEPLOYMENT & TESTING DOCS
In /docs include:
- Local dev quickstart (HTTPS dev cert, running add-in locally, sideloading)
- Cloud deploy guide (deploy API + DB + configure app registrations + update manifest URLs)
- Sovereign/Government notes: Centralized Deployment requirements and alternatives (SharePoint catalog / trusted catalogs).
- Troubleshooting section: common SSO issues, token validation errors, CORS, mixed-content HTTPS problems.
- A short architecture diagram (ASCII or Mermaid) showing: Word -> Taskpane -> getAccessToken -> API -> Azure SQL.

IMPLEMENTATION DETAILS TO INCLUDE
- Provide all code files needed (no placeholders like “TODO implement”).
- Ensure build/run instructions are accurate and scripts are runnable.
- Ensure the add-in can be launched in Word desktop and Word web (where possible).
- Use modern tooling and clear comments.

DELIVERABLES
Generate the full repo with:
- All source code under the folders described
- manifest.xml (or unified manifest if you choose; document which you used)
- Working snippet insertion for text and images
- /docs/Admin-Setup.md and /docs/Developer-Setup.md and /docs/Deployment.md
- /infra/main.bicep plus any modules
- /infra/deploy.ps1 or deploy.sh using az CLI
- sample SQL scripts and/or EF migrations
- a short README.md at root with step-by-step instructions

QUALITY BAR
This is a demo for an enterprise/government customer scenario. Security and documentation must be clear. No secrets in client code. Use HTTPS everywhere. Validate tokens on server. Use managed identity to Azure SQL. Provide fallback notes if Office SSO fails (optional MSAL fallback), but keep the main path Office SSO.

Now create the repository structure and output all files with their full contents. When you reference URLs or IDs that must be replaced, clearly mark them and list them in a "Replace These Values" section. 
``

## GCC (US Government Community Cloud) constraints addendum (paste into the Copilot prompt)

**Target environment:** US Government Community Cloud (GCC) — *not* GCC High and *not* DoD. [1](https://stackoverflow.com/questions/71995699/office-add-ins-with-blazor-wasm)

### 1) Acquisition + deployment constraints (GCC)
- **No public Microsoft 365 / Copilot store:** End users **can’t acquire** Office Add-ins directly from the public store, and admins **can’t deploy** Office Add-ins from the public store in the Admin Portal in sovereign/government clouds (includes GCC). [1](https://stackoverflow.com/questions/71995699/office-add-ins-with-blazor-wasm)  
- **Manifest-based deployment is expected:** The Office Add-in **manifest is obtained from the developer/partner and ingested for deployment** via the Admin Portal; **centralized deployment outside of the store is supported**. [1](https://stackoverflow.com/questions/71995699/office-add-ins-with-blazor-wasm)[2](https://stackoverflow.com/questions/64138534/blazor-as-a-office-add-in)  
- **Integrated apps portal limitation:** Microsoft notes that customers in sovereign/government clouds **don’t have access to the integrated apps portal**, so deployment planning should rely on the supported admin deployment guidance for these clouds. [3](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/publish/government-cloud-guidance.md?plain=1)[2](https://stackoverflow.com/questions/64138534/blazor-as-a-office-add-in)  

### 2) Cloud boundary / network constraints (GCC)
- **Assume “inside boundary” first:** Required resources/services should be available **inside the cloud boundary**, or the customer’s network admin must explicitly allow access to resources **outside** the boundary. [1](https://stackoverflow.com/questions/71995699/office-add-ins-with-blazor-wasm)  
- **Data handling & vetting:** The add-in’s accessed resources must conform to GCC requirements and any tenant-specific requirements around **transfer, management, storage of sensitive data**, and **code/resource security & access vetting**. [1](https://stackoverflow.com/questions/71995699/office-add-ins-with-blazor-wasm)  

### 3) Centralized Deployment prerequisites to call out in admin/runbook (GCC)
- Users must have **Exchange Online** and **active Exchange Online mailboxes** for centralized deployment eligibility. [2](https://stackoverflow.com/questions/64138534/blazor-as-a-office-add-in)  
- The directory must be **in or federated to Microsoft Entra ID**. [2](https://stackoverflow.com/questions/64138534/blazor-as-a-office-add-in)  
- Supported licenses include **Office 365 Government (G3/G5)** and **Microsoft 365 Government (G3/G5)** (among others). [2](https://stackoverflow.com/questions/64138534/blazor-as-a-office-add-in)  
- Centralized deployment is **not supported** for **on-premises Exchange mailboxes**. [2](https://stackoverflow.com/questions/64138534/blazor-as-a-office-add-in)  

### 4) Approved deployment options to document for GCC
- **Primary:** Centralized deployment using the Admin Center with a **manifest** (not store-based). [2](https://stackoverflow.com/questions/64138534/blazor-as-a-office-add-in)[1](https://stackoverflow.com/questions/71995699/office-add-ins-with-blazor-wasm)  
- **Fallback / special cases:** SharePoint app catalog can be used to host manifests, but it has limitations (for example, Outlook add-ins aren’t supported by app catalogs; and there are feature limitations tied to the manifest). [4](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)  
- **Dev/test only:** Network share “trusted catalog” sideloading is described as a testing approach and involves configuring trusted catalogs (manual or registry-based). [5](https://microsoft-my.sharepoint.com/personal/jobyford_microsoft_com/Documents/Recordings/New-Ish%20To%20Microsoft%20%28Optional%29-20260306_070240-Meeting%20Recording.mp4?web=1)  

### 5) Implementation guidance to force in generated solution (GCC)
- The generated demo must **not rely on Microsoft Marketplace/store** for installation, and must include:  
  - a **manifest-only** deployment path for admins,  
  - a **GCC-friendly** admin runbook (Centralized Deployment prerequisites + steps),  
  - and explicit “inside boundary / approved egress” notes for all endpoints. [1](https://stackoverflow.com/questions/71995699/office-add-ins-with-blazor-wasm)[2](https://stackoverflow.com/questions/64138534/blazor-as-a-office-add-in)[3](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/publish/government-cloud-guidance.md?plain=1)
