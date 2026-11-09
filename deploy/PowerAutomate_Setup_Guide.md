# Power Automate & SharePoint Setup Guide (Option 4 M365)
**System:** UHBFT Pharmacy Near Miss Reporting & Learning System  
**Document Date:** 9 November 2026

---

## 1. Architecture Overview

To maintain the sub-10ms barcode scanning and type-ahead search across `101,478` GTINs + `10,125` CMM items without triggering the PowerApps 2,000-row delegation ceiling, the app uses a **Hybrid Edge-to-Cloud pattern**:
* **Frontend:** Fast HTML5 bench UI (`Near-Miss-Logger.html` + `nearmiss_drug_index.js`) running on checking bench PCs.
* **Backend:** Microsoft 365 Power Automate Cloud Flow receiving submissions via HTTP POST and inserting records directly into the governed SharePoint List **`Pharmacy Near Miss Log`**.

---

## 2. Step 1: Provision SharePoint List

1. Navigate to the Trust Pharmacy SharePoint Team Site (e.g. `https://uhbnhs.sharepoint.com/teams/PharmacySafety`).
2. Click **New** -> **List** -> **From JSON** (or create matching columns specified in `SharePoint_Schema.json`).
3. Ensure the list title is set to **`Pharmacy Near Miss Log`**.

---

## 3. Step 2: Create Power Automate Cloud Flow

1. Open Power Automate (`https://make.powerautomate.com/`) under the Trust tenant.
2. Click **Create** -> **Automated Cloud Flow** (or Instant Flow with HTTP trigger).
3. **Trigger:** Select **When an HTTP request is received**.
   * **Request Body JSON Schema:** Paste the contents of `SharePoint_Schema.json` or sample canonical payload from `PLAN_Near_Miss_Reporting_App.md`.
   * **Who can trigger:** `Any user in my tenant` (Restricted to UHBFT accounts).
4. **Action:** Add **Create item** (SharePoint connector):
   * **Site Address:** Select the Pharmacy Safety SharePoint site.
   * **List Name:** `Pharmacy Near Miss Log`.
   * **Field Mappings:**
     * `Title` -> `triggerBody()?['eventId']`
     * `SubmittedAt` -> `triggerBody()?['submittedAtIso']`
     * `Site` -> `triggerBody()?['site']`
     * `NearMissType` -> `triggerBody()?['nearMissType']`
     * `ErrorDrugName` -> `triggerBody()?['errorDrug']?['displayName']`
     * `ErrorDrugGTIN14` -> `triggerBody()?['errorDrug']?['scannedGtin14']`
     * `ErrorDrugBatch` -> `triggerBody()?['errorDrug']?['batchNumber']`
     * `ErrorDrugExpiry` -> `triggerBody()?['errorDrug']?['expiryDate']`
     * `IntendedDrugName` -> `triggerBody()?['intendedDrug']?['displayName']`
     * `LearningNote` -> `triggerBody()?['learningNote']`
     * *(Map remaining fields as per schema)*.
5. **Action:** Add **Response** (HTTP 200 with `{"success": true, "id": triggerBody()?['eventId']}`).
6. Save the Flow and copy the generated **HTTP POST URL**.

---

## 4. Step 3: Activate in `nearmiss_config.js`

Open `NearMiss\nearmiss_config.js` and set:
```javascript
window.NEARMISS_CONFIG = {
  storageMode: 'sharepoint-webhook',
  sharepointWebhookUrl: 'https://prod-XX.westeurope.logic.azure.com:443/workflows/...',
  ...
};
```
Checkers will instantly write to Trust SharePoint Lists with zero change in user experience.
