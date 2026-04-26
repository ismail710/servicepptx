# ProjectHub PPTX Download — No Azure Setup Guide

## What this does
A Power Automate flow triggered from Power Apps that:
1. Receives the project JSON payload from Power Apps
2. Calls a small Python service to fill `template1.pptx` with the project data
3. Returns the filled PPTX as base64 → Power Apps triggers an immediate browser download

No Azure Functions, no email attachment required.

---

## Architecture

```
Power Apps button
      │  (projectPayloadJson, serviceUrl)
      ▼
Power Automate flow  (ProjectHubPptxDownload.zip)
      │  POST /generate  { project: {...} }
      ▼
pptx_service.py  (hosted on Render.com FREE or run locally)
      │  fills template1.pptx → returns { pptxBase64, fileName }
      ▼
Power Automate → responds to Power Apps
      │
      ▼
Power Apps  →  Launch("data:application/...;base64," & pptxBase64)
                        ↓
                  browser downloads .pptx
```

---

## Step 1 — Deploy the Python service (FREE, no Azure)

### Option A — Render.com (recommended, always-on free tier)

1. Push this repo to GitHub (or GitLab / Bitbucket)
2. Go to **render.com** → sign up (free) → **New → Web Service**
3. Connect your repo
4. Render auto-detects `render.yaml` and configures everything
5. Click **Create Web Service**
6. Your deployed URL: `https://servicepptx.onrender.com`
7. Test it: `GET https://servicepptx.onrender.com/api/health` → `{"status":"ok"}`

> **Note:** Free tier spins down after 15 min of inactivity (cold start ~30 s).
> Upgrade to Starter ($7/month) for always-on.

### Option B — Run locally (for testing)

```powershell
# Inside the project folder:
.\.venv\Scripts\python.exe pptx_service.py
# Service listens on http://localhost:7771
```

To expose locally to Power Automate use **ngrok**:
```powershell
ngrok http 7771
# Copy the https://xxxxx.ngrok.io URL
```

### Option C — Railway.app (alternative free host)

1. Go to **railway.app** → New Project → Deploy from GitHub Repo
2. Set environment variable `PORT=8080`
3. Add `TEMPLATE_PATH=template1.pptx`
4. Copy the generated URL

---

## Step 2 — Import the Power Automate solution

1. Go to **make.powerautomate.com**
2. **Solutions → Import** → upload `ProjectHubPptxDownload.zip`
3. Import completes with no connection references needed (uses HTTP, no connectors)
4. Open the imported flow **GeneratePptxForDownload-ProjectHub** and confirm it is **On**

---

## Step 3 — Connect Power Apps to the flow

In Power Apps Studio, select the download button and add the flow:

**OnSelect** of your download button:

```powerapps
// 1. Build the JSON payload from the current project
Set(
    varPptxPayload,
    JSON({
        projectNumber:     Text(CurrentItem.'Numéro projet'),
        projectTitle:      CurrentItem.Libellé,
        projectLead:       CurrentItem.'Chef de projet'.DisplayName,
        piUm6p:            CurrentItem.'PI UM6P',
        startDate:         Text(CurrentItem.'Date début', "dd/mm/yyyy"),
        endDate:           Text(CurrentItem.'Date fin',   "dd/mm/yyyy"),
        projectDescription:CurrentItem.Commentaires,
        strategicAxis:     CurrentItem.'Axe stratégique',
        budget:            Text(CurrentItem.Budget),
        budgetType:        CurrentItem.'Type budget',
        finality:          CurrentItem.Finalité,
        deliverablesCsv:   Concat(
                               Filter([@Livrables], Projet.Projet = CurrentItem.Projet),
                               Livrable, " | "),
        risksAndAlertsCsv: CurrentItem.Risques,
        doneCsv:           Concat(
                               Filter([@Livrables], Projet.Projet = CurrentItem.Projet
                                   && !IsBlank('Date reçue')),
                               Livrable, " | "),
        inProgressCsv:     Concat(
                               Filter([@Livrables], Projet.Projet = CurrentItem.Projet
                                   && IsBlank('Date reçue') && 'Date prévue' <= Today()),
                               Livrable, " | "),
        plannedCsv:        Concat(
                               Filter([@Livrables], Projet.Projet = CurrentItem.Projet
                                   && IsBlank('Date reçue') && 'Date prévue' > Today()),
                               Livrable, " | "),
        capexCommitments:  Text(CurrentItem.CAPEX_Engagements),
        capexExpenses:     Text(CurrentItem.CAPEX_Dépenses),
        opexCommitments:   Text(CurrentItem.OPEX_Engagements),
        opexExpenses:      Text(CurrentItem.OPEX_Dépenses)
    })
);

// 2. Call the flow  ← replace "GeneratePptxForDownload_ProjectHub" with actual flow name
Set(
    varPptxResult,
    GeneratePptxForDownload_ProjectHub.Run(
        varPptxPayload,
        "https://servicepptx.onrender.com"   // ← live Render service URL
    )
);

// 3. Trigger browser download
If(
    !IsBlank(varPptxResult.pptxbase64),
    Launch(
        "data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,"
            & varPptxResult.pptxbase64
    ),
    Notify("Erreur lors de la génération du PPTX", NotificationType.Error)
)
```

---

## Template placeholders (template1.pptx)

| Placeholder            | Field (JSON key)       |
|------------------------|------------------------|
| `{{Project_Number}}`   | `projectNumber`        |
| `{{Project_Title}}`    | `projectTitle`         |
| `{{Project_Lead}}`     | `projectLead`          |
| `{{PI_UM6P}}`          | `piUm6p`               |
| `{{Starting_Date}}`    | `startDate`            |
| `{{Closing_Date}}`     | `endDate`              |
| `{{Budget_Amount}}`    | `budget`               |
| `{{Deliverables_List}}`| `deliverablesCsv`      |
| `{{Risk_Alert}}`       | `risksAndAlertsCsv`    |
| `{{Tasks_Done}}`       | `doneCsv`              |
| `{{Tasks_InProgress}}` | `inProgressCsv`        |
| `{{Tasks_NextSteps}}`  | `plannedCsv`           |
| `{{CAPEX_Commitments}}`| `capexCommitments`     |
| `{{CAPEX_Expenses}}`   | `capexExpenses`        |
| `{{OPEX_Commitments}}` | `opexCommitments`      |
| `{{OPEX_Expenses}}`    | `opexExpenses`         |

---

## Files created

| File | Purpose |
|------|---------|
| `pptx_service.py` | Python HTTP service — fills template and returns base64 |
| `requirements.txt` | `python-pptx` + `lxml` |
| `render.yaml` | Render.com free deployment config |
| `Procfile` | Alternative for Railway/Heroku |
| `ProjectHubPptxDownload.zip` | Importable Power Automate solution |
| `template1.pptx` | PowerPoint template with `{{placeholders}}` |
