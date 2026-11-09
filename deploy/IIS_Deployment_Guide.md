# IIS Intranet Deployment Guide (Option 2 Central Server)
**System:** UHBFT Pharmacy Near Miss Reporting & Learning System  
**Document Date:** 9 November 2026

---

## 1. Overview

Option 2 hosts the entire static bundle (`Near-Miss-Logger.html`, `nearmiss_drug_index.js`, `nearmiss_config.js`, `Near-Miss-Analytics-Dashboard.html`) on an internal Trust IIS Web Server (e.g. `https://pharmacy-intranet.uhb.nhs.uk/nearmiss/`).

---

## 2. Directory Structure on IIS Server

```
C:\inetpub\wwwroot\PharmacyNearMiss\
├── Near-Miss-Logger.html
├── Near-Miss-Analytics-Dashboard.html
├── nearmiss_drug_index.js
├── nearmiss_config.js
├── web.config
└── api\
    ├── nearmiss.ashx (or .NET Core Web API)
    └── data\
        └── events\
```

---

## 3. Minimal `web.config`

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <system.webServer>
    <staticContent>
      <remove fileExtension=".json" />
      <mimeMap fileExtension=".json" mimeType="application/json; charset=utf-8" />
    </staticContent>
    <httpProtocol>
      <customHeaders>
        <add name="X-Content-Type-Options" value="nosniff" />
        <add name="X-Frame-Options" value="SAMEORIGIN" />
      </customHeaders>
    </httpProtocol>
  </system.webServer>
</configuration>
```

---

## 4. Activating in `nearmiss_config.js`

```javascript
window.NEARMISS_CONFIG = {
  storageMode: 'iis-api',
  iisEndpoint: 'https://pharmacy-intranet.uhb.nhs.uk/api/nearmiss',
  ...
};
```
Checkers will instantly write to Trust IIS servers with zero change in user experience.
