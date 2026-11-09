/* ==============================================================================
   UHBFT Pharmacy Near Miss Reporting & Learning System
   Configuration & Storage Provider Settings
   ============================================================================== */
window.NEARMISS_CONFIG = {
  // Storage Provider: 'local-bridge' (default Stage 1), 'iis-api' (Option 2), 'sharepoint-webhook' (Option 4)
  storageMode: 'local-bridge',

  // Stage 1: Local / Network Share Micro-Server Bridge
  localBridgeUrl: 'http://localhost:8088',

  // Stage 2A: IIS Intranet API Endpoint (Option 2)
  iisEndpoint: 'https://pharmacy-intranet.uhb.nhs.uk/api/nearmiss',

  // Stage 2B: Microsoft 365 Power Automate / SharePoint Webhook (Option 4)
  sharepointWebhookUrl: '',

  // Offline LocalStorage Queue Key
  offlineQueueKey: 'nearmiss_offline_queue',

  // Undo window duration in seconds
  undoWindowSeconds: 5,

  // Trust Hospital Sites
  sites: [
    'Queen Elizabeth',
    'Heartlands',
    'Good Hope',
    'Solihull',
    'Other'
  ],

  // Satellite Units
  satelliteUnits: [
    'GHH Satellite Unit',
    'SH Satellite Unit',
    'Other Satellite'
  ]
};
