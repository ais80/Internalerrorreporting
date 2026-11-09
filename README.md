# Internal Error Reporting & Near Miss Safety Suite (PWA)

A high-speed, sub-8-second checking bench and mobile reporting suite for NHS hospital pharmacy dispensaries and satellite units. Runs 100% offline or integrated with hospital intranet servers and Microsoft 365.

---

## Key Features

- **Progressive Web App (PWA)**: Works on desktop checking benches, tablets, and mobile phones (iOS & Android) with zero app store install required.
- **Dual Barcode Scanner Engine**:
  - **Hardware Keyboard Wedge**: Captures rapid streams from USB and cordless scanners (Zebra, Honeywell, Datalogic) without needing cursor focus.
  - **Mobile Camera Imager**: Real-time camera viewfinder with hardware `BarcodeDetector` API + offline ZXing fallback for scanning pack barcodes directly from mobile devices.
- **GS1 2D DataMatrix & 1D EAN Decoding**:
  - Extracts GTIN-14, Expiry Date (`AI 17`), and Batch Number (`AI 10`).
  - Automatically detects and flags expired or short-dated packs.
- **Formulary & dm+d Intelligence**:
  - Resolves `GTIN -> AMPP -> AMP -> VMP -> VTM` in `< 5ms`.
  - Highlights Narrow Therapeutic Index (NTI) brand-critical drugs.
  - Sibling drug intelligence computes instant 1-tap chips for same-drug strengths, formulations, and brands.
- **Information Governance (IG) Safeguard**:
  - **Zero Patient Identifiable Data (PID)**.
  - Client-side regex interceptor blocking accidental entry of 10-digit NHS numbers or hospital MRNs.
- **Multi-Site Pre-Configuration**:
  - Sticky local profile for checking benches.
  - Site QR code poster routing (`?site=GHH`, `?site=BHH`, `?site=QEHB`, `?site=SH`).

---

## File Structure

```
├── Near-Miss-Logger.html               # Main checking bench & mobile PWA logger
├── Near-Miss-Analytics-Dashboard.html  # Governance analytics, LASA matrices & heatmaps
├── nearmiss_config.js                  # Site, satellite, and storage mode settings
├── nearmiss_drug_index.js              # Compiled offline dm+d & stock dictionary
├── nearmiss_shared_data.js             # Shared reporting definitions & taxonomy
├── zxing.min.js                        # Offline mobile camera barcode decoder
├── manifest.json                       # Progressive Web App manifest
├── nearmiss-server.ps1                 # Lightweight local micro-bridge server
├── Sync-NearMissMaster.ps1             # Consolidates atomic events into Master Excel & CSV
└── Build-NearMissDrugIndex.ps1         # Compiles dm+d TRUD + stock files into dictionary
```

---

## Getting Started

### 1. Launching on a Checking Bench (Desktop)
Run the local micro-bridge server:
```powershell
.\nearmiss-server.ps1
```
Open in Microsoft Edge / Chrome:
```
http://localhost:8088/Near-Miss-Logger.html
```

### 2. Using on Mobile Phones / Tablets (PWA)
1. Ensure the phone is connected to the Trust / local Wi-Fi.
2. Open the mobile URL displayed in the console (e.g. `http://<IP>:8088/Near-Miss-Logger.html?site=GHH`).
3. Tap **"Share" -> "Add to Home Screen"** on iOS or **"Install App"** on Android to run full-screen.
4. Tap the **📷 Camera Scan** button to scan medicine pack barcodes.
