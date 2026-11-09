# ==============================================================================
# Build-NearMissDrugIndex.ps1
# UHBFT Secondary Care Pharmacy Near Miss Reporting & Learning System
# Pre-compiled GTIN + dm+d + CMM Drug Index Builder (Phase 1)
# ==============================================================================
[CmdletBinding()]
param(
    [string]$DmdDir = "",
    [string]$CmmFile = "",
    [string]$OutJs = "",
    [string]$OutMetrics = ""
)

# Auto-detect Dashboard directory whether running from NearMiss\ or parent root
$dashboardDir = ""
$candidates = @(
    "$PSScriptRoot\..\..\..\Dot Prescriptions - weekly check\_Dashboard",
    "$PSScriptRoot\..\..\Dot Prescriptions - weekly check\_Dashboard",
    "$PSScriptRoot\Dot Prescriptions - weekly check\_Dashboard"
)
foreach ($c in $candidates) {
    if (Test-Path $c) {
        $dashboardDir = (Get-Item $c).FullName
        break
    }
}

if (-not $dashboardDir) {
    Write-Error "Could not locate 'Dot Prescriptions - weekly check\_Dashboard' directory."
    exit 1
}

if (-not $DmdDir) { $DmdDir = "$dashboardDir\dmdDataLoader\csv" }
if (-not $CmmFile) { $CmmFile = "$dashboardDir\DM+D files\CMM_Drug_file_dmd_matched.csv" }
if (-not $OutJs) { $OutJs = "$PSScriptRoot\nearmiss_drug_index.js" }
if (-not $OutMetrics) { $OutMetrics = "$PSScriptRoot\drug_index_metrics.json" }

# Resolve absolute paths
$DmdDir = [System.IO.Path]::GetFullPath($DmdDir)
$CmmFile = [System.IO.Path]::GetFullPath($CmmFile)
$OutJs = [System.IO.Path]::GetFullPath($OutJs)
$OutMetrics = [System.IO.Path]::GetFullPath($OutMetrics)

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " UHBFT Pharmacy Near Miss - Compiling Drug & Barcode Index Engine (Phase 1)     " -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " dm+d Directory : $DmdDir"
Write-Host " CMM File       : $CmmFile"
Write-Host " Target Output  : $OutJs"
Write-Host " Metrics File   : $OutMetrics"
Write-Host ""

if (-not (Test-Path $DmdDir)) {
    Write-Error "dm+d directory not found: $DmdDir"
    exit 1
}
if (-not (Test-Path $CmmFile)) {
    Write-Error "CMM matched file not found: $CmmFile"
    exit 1
}

$csharpCode = @"
using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

public class NearMissIndexCompiler {
    public static string Execute(string dmdDir, string cmmPath, string outJsPath, string outMetricsPath) {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var ntiRegex = new Regex("tacrolimus|ciclosporin|cyclosporine|lithium|diltiazem|theophylline|carbamazepine|phenytoin|valproate|amiodarone|flecainide|clozapine|digoxin|mesalazine|methotrexate|adalimumab|infliximab|etanercept|rituximab|trastuzumab|insulin|filgrastim", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // 1. VTM Dictionary: VTMID -> VTM Name
        var vtmMap = new Dictionary<string, string>();
        using (var reader = File.OpenText(Path.Combine(dmdDir, "f_vtm.csv"))) {
            string line;
            while ((line = reader.ReadLine()) != null) {
                var p = line.Split('|');
                if (p.Length >= 3) vtmMap[p[0]] = p[2];
            }
        }

        // 2. VMP Dictionaries: VMPID -> VTMID, Name
        var vmpVtm = new Dictionary<string, string>();
        var vmpNm = new Dictionary<string, string>();
        using (var reader = File.OpenText(Path.Combine(dmdDir, "f_vmp_VmpType.csv"))) {
            string line;
            while ((line = reader.ReadLine()) != null) {
                var p = line.Split('|');
                if (p.Length > 5 && string.IsNullOrEmpty(p[4])) { // active VMP only
                    vmpVtm[p[0]] = p[3];
                    vmpNm[p[0]] = p[5];
                }
            }
        }

        // 3. AMP Dictionaries: AMPID -> VMPID, Name, Description (with manufacturer)
        var ampVmp = new Dictionary<string, string>();
        var ampNm = new Dictionary<string, string>();
        var ampDesc = new Dictionary<string, string>();
        using (var reader = File.OpenText(Path.Combine(dmdDir, "f_amp_AmpType.csv"))) {
            string line;
            while ((line = reader.ReadLine()) != null) {
                var p = line.Split('|');
                if (p.Length > 5) {
                    ampVmp[p[0]] = p[2];
                    ampNm[p[0]] = p[3];
                    ampDesc[p[0]] = p[5];
                }
            }
        }

        // 4. AMPP Dictionaries: AMPPID -> AMPID, Name, DiscontinuedFlag
        var amppAmp = new Dictionary<string, string>();
        var amppNm = new Dictionary<string, string>();
        var amppDisc = new Dictionary<string, int>();
        using (var reader = File.OpenText(Path.Combine(dmdDir, "f_ampp_AmppType.csv"))) {
            string line;
            while ((line = reader.ReadLine()) != null) {
                var p = line.Split('|');
                if (p.Length > 5) {
                    amppAmp[p[0]] = p[5];
                    amppNm[p[0]] = p[2];
                    amppDisc[p[0]] = (p.Length > 9 && p[9] == "0001") ? 1 : 0;
                }
            }
        }

        // 5. CMM Stock (All 10,125 lines grouped by CMM_lnkdid)
        var concepts = new List<string>();
        var vmpToConcept = new Dictionary<string, int>();
        var ampToConcept = new Dictionary<string, int>();
        var cmmCodeToConcept = new Dictionary<string, int>();
        var vtmSiblings = new Dictionary<string, List<int>>();

        int cmmLineCount = 0;
        using (var reader = File.OpenText(cmmPath)) {
            string headerLine = reader.ReadLine();
            string line;
            var cmmGroup = new Dictionary<string, string[]>();
            while ((line = reader.ReadLine()) != null) {
                cmmLineCount++;
                var cols = ParseCsvRow(line);
                if (cols.Length < 18) continue;
                string lnkdid = cols[0];
                if (!cmmGroup.ContainsKey(lnkdid)) {
                    cmmGroup[lnkdid] = cols;
                }
            }

            foreach (var kvp in cmmGroup) {
                var cols = kvp.Value;
                string lnkdid = cols[0];
                string drugName = cols[1];
                string strength = cols[2];
                string form = cols[3];
                string fullName = cols[4];
                if (string.IsNullOrWhiteSpace(fullName)) {
                    fullName = (drugName + " " + strength + " " + form).Trim();
                }
                string vtmCode = cols[8];
                string vtmName = cols[9];
                if (string.IsNullOrWhiteSpace(vtmName) && !string.IsNullOrWhiteSpace(vtmCode) && vtmMap.ContainsKey(vtmCode)) {
                    vtmName = vtmMap[vtmCode];
                }
                string vmpCode = cols[14];
                string vmpName = cols[15];
                if (string.IsNullOrWhiteSpace(vmpName) && !string.IsNullOrWhiteSpace(vmpCode) && vmpNm.ContainsKey(vmpCode)) {
                    vmpName = vmpNm[vmpCode];
                }
                string ampCode = cols[16];
                string ampName = cols[17];
                if (string.IsNullOrWhiteSpace(ampName) && !string.IsNullOrWhiteSpace(ampCode) && ampDesc.ContainsKey(ampCode)) {
                    ampName = ampDesc[ampCode];
                }

                int isNti = ntiRegex.IsMatch(fullName + " " + drugName + " " + vtmName + " " + vmpName + " " + ampName) ? 1 : 0;
                int cIdx = concepts.Count;

                string dgOpt = string.Equals(drugName.Trim(), vtmName.Trim(), StringComparison.OrdinalIgnoreCase) ? "" : drugName;
                string vpnOpt = string.Equals(fullName.Trim(), vmpName.Trim(), StringComparison.OrdinalIgnoreCase) ? "" : vmpName;

                var sb = new StringBuilder();
                sb.Append("[\"").Append(Escape(lnkdid)).Append("\",")
                  .Append("\"").Append(Escape(fullName)).Append("\",")
                  .Append("\"").Append(Escape(dgOpt)).Append("\",")
                  .Append("\"").Append(Escape(strength)).Append("\",")
                  .Append("\"").Append(Escape(form)).Append("\",")
                  .Append("\"").Append(Escape(lnkdid)).Append("\",")
                  .Append("\"").Append(Escape(vtmCode)).Append("\",")
                  .Append("\"").Append(Escape(vtmName)).Append("\",")
                  .Append("\"").Append(Escape(vmpCode)).Append("\",")
                  .Append("\"").Append(Escape(vpnOpt)).Append("\",")
                  .Append("\"").Append(Escape(ampCode)).Append("\",")
                  .Append("\"").Append(Escape(ampName)).Append("\",")
                  .Append(isNti).Append(",1]");
                concepts.Add(sb.ToString());

                cmmCodeToConcept[lnkdid] = cIdx;
                if (!string.IsNullOrWhiteSpace(vmpCode) && !vmpToConcept.ContainsKey(vmpCode)) {
                    vmpToConcept[vmpCode] = cIdx;
                }
                if (!string.IsNullOrWhiteSpace(ampCode) && !ampToConcept.ContainsKey(ampCode)) {
                    ampToConcept[ampCode] = cIdx;
                }

                if (!string.IsNullOrWhiteSpace(vtmCode)) {
                    if (!vtmSiblings.ContainsKey(vtmCode)) vtmSiblings[vtmCode] = new List<int>();
                    vtmSiblings[vtmCode].Add(cIdx);
                }
            }
        }

        // 6. National VMPs not in CMM
        int nationalVmpAdded = 0;
        foreach (var kvp in vmpNm) {
            string vmpId = kvp.Key;
            if (!vmpToConcept.ContainsKey(vmpId)) {
                string vtmId = vmpVtm.ContainsKey(vmpId) ? vmpVtm[vmpId] : "";
                string vtmName = (!string.IsNullOrWhiteSpace(vtmId) && vtmMap.ContainsKey(vtmId)) ? vtmMap[vtmId] : "";
                string vmpName = kvp.Value;
                int isNti = ntiRegex.IsMatch(vmpName + " " + vtmName) ? 1 : 0;

                int cIdx = concepts.Count;
                var sb = new StringBuilder();
                sb.Append("[\"").Append(Escape(vmpId)).Append("\",")
                  .Append("\"").Append(Escape(vmpName)).Append("\",")
                  .Append("\"\",") // dg == vn
                  .Append("\"\",\"\",")
                  .Append("\"\",") // cm == empty
                  .Append("\"").Append(Escape(vtmId)).Append("\",")
                  .Append("\"").Append(Escape(vtmName)).Append("\",")
                  .Append("\"").Append(Escape(vmpId)).Append("\",")
                  .Append("\"\",") // vpn == nm
                  .Append("\"\",\"\",")
                  .Append(isNti).Append(",0]");
                concepts.Add(sb.ToString());

                vmpToConcept[vmpId] = cIdx;
                nationalVmpAdded++;

                if (!string.IsNullOrWhiteSpace(vtmId)) {
                    if (!vtmSiblings.ContainsKey(vtmId)) vtmSiblings[vtmId] = new List<int>();
                    vtmSiblings[vtmId].Add(cIdx);
                }
            }
        }

        // 7. Branded AMPs: Include NTI brands
        int brandAmpAdded = 0;
        foreach (var kvp in ampNm) {
            string ampId = kvp.Key;
            if (!ampToConcept.ContainsKey(ampId)) {
                string vmpId = ampVmp.ContainsKey(ampId) ? ampVmp[ampId] : "";
                string aNm = kvp.Value;
                string vNm = (!string.IsNullOrWhiteSpace(vmpId) && vmpNm.ContainsKey(vmpId)) ? vmpNm[vmpId] : "";

                if (!string.IsNullOrWhiteSpace(vNm) && !string.Equals(aNm.Trim(), vNm.Trim(), StringComparison.OrdinalIgnoreCase)) {
                    if (ntiRegex.IsMatch(aNm)) {
                        string aDesc = ampDesc.ContainsKey(ampId) ? ampDesc[ampId] : aNm;
                        string vtmId = (!string.IsNullOrWhiteSpace(vmpId) && vmpVtm.ContainsKey(vmpId)) ? vmpVtm[vmpId] : "";
                        string vtmName = (!string.IsNullOrWhiteSpace(vtmId) && vtmMap.ContainsKey(vtmId)) ? vtmMap[vtmId] : "";

                        int cIdx = concepts.Count;
                        var sb = new StringBuilder();
                        sb.Append("[\"").Append(Escape(ampId)).Append("\",")
                          .Append("\"").Append(Escape(aNm)).Append("\",")
                          .Append("\"\",")
                          .Append("\"\",\"\",")
                          .Append("\"\",")
                          .Append("\"").Append(Escape(vtmId)).Append("\",")
                          .Append("\"").Append(Escape(vtmName)).Append("\",")
                          .Append("\"").Append(Escape(vmpId)).Append("\",")
                          .Append("\"").Append(Escape(vNm)).Append("\",")
                          .Append("\"").Append(Escape(ampId)).Append("\",")
                          .Append("\"").Append(Escape(aDesc)).Append("\",")
                          .Append("1,0]");
                        concepts.Add(sb.ToString());

                        ampToConcept[ampId] = cIdx;
                        brandAmpAdded++;

                        if (!string.IsNullOrWhiteSpace(vtmId)) {
                            if (!vtmSiblings.ContainsKey(vtmId)) vtmSiblings[vtmId] = new List<int>();
                            vtmSiblings[vtmId].Add(cIdx);
                        }
                    }
                }
            }
        }

        // 8. Packs Suffix Table & GTIN Map
        var packList = new List<string>();
        var packIndexMap = new Dictionary<string, int>();
        var gtinMapLines = new List<string>();
        int gtinTotal = 0;
        int gtinResolved = 0;

        using (var reader = File.OpenText(Path.Combine(dmdDir, "f_gtin.csv"))) {
            string line;
            while ((line = reader.ReadLine()) != null) {
                var p = line.Split('|');
                gtinTotal++;
                string amppId = p[0];
                string gtin14 = p[1].Trim().PadLeft(14, '0');
                int isHist = (p.Length > 3 && !string.IsNullOrWhiteSpace(p[3])) ? 1 : 0;

                string ampId = "";
                string vmpId = "";
                string packDesc = "";
                int isDisc = 0;
                int cIdx = -1;

                if (amppAmp.ContainsKey(amppId)) {
                    ampId = amppAmp[amppId];
                    packDesc = amppNm.ContainsKey(amppId) ? amppNm[amppId] : "";
                    isDisc = amppDisc.ContainsKey(amppId) ? amppDisc[amppId] : 0;

                    if (ampVmp.ContainsKey(ampId)) {
                        vmpId = ampVmp[ampId];
                    }

                    if (ampToConcept.ContainsKey(ampId)) {
                        cIdx = ampToConcept[ampId];
                    } else if (!string.IsNullOrWhiteSpace(vmpId) && vmpToConcept.ContainsKey(vmpId)) {
                        cIdx = vmpToConcept[vmpId];
                    }
                }

                if (cIdx >= 0) gtinResolved++;

                string packSuffix = packDesc;
                if (!string.IsNullOrEmpty(ampId) && ampDesc.ContainsKey(ampId)) {
                    string aDesc = ampDesc[ampId];
                    if (packDesc.StartsWith(aDesc)) {
                        packSuffix = packDesc.Substring(aDesc.Length).Trim();
                    }
                }

                int pIdx = -1;
                if (!string.IsNullOrEmpty(packSuffix)) {
                    if (!packIndexMap.TryGetValue(packSuffix, out pIdx)) {
                        pIdx = packList.Count;
                        packList.Add(packSuffix);
                        packIndexMap[packSuffix] = pIdx;
                    }
                }

                gtinMapLines.Add("\"" + gtin14 + "\":[\"" + amppId + "\",\"" + ampId + "\"," + pIdx + "," + isHist + "," + isDisc + "," + cIdx + "]");
            }
        }

        // 9. Write nearmiss_drug_index.js
        using (var writer = new StreamWriter(outJsPath, false, Encoding.UTF8)) {
            writer.Write("/* UHBFT Pharmacy Near Miss Drug Index - Auto-generated by Build-NearMissDrugIndex.ps1 */\n");
            writer.Write("window.NEARMISS_DRUG_DB = {\n");
            writer.Write("  generatedAt: \"" + DateTime.UtcNow.ToString("o") + "\",\n");
            writer.Write("  gtinTotal: " + gtinTotal + ",\n");
            writer.Write("  conceptTotal: " + concepts.Count + ",\n");

            // Packs array
            writer.Write("  packs: [");
            for (int i = 0; i < packList.Count; i++) {
                if (i > 0) writer.Write(",");
                writer.Write("\"" + Escape(packList[i]) + "\"");
            }
            writer.Write("],\n");

            // Concepts array
            writer.Write("  concepts: [\n");
            for (int i = 0; i < concepts.Count; i++) {
                if (i > 0) writer.Write(",\n");
                writer.Write(concepts[i]);
            }
            writer.Write("\n  ],\n");

            // VTM Siblings
            writer.Write("  vtmSiblings: {\n");
            bool firstSibling = true;
            foreach (var kvp in vtmSiblings) {
                if (!firstSibling) writer.Write(",\n"); else firstSibling = false;
                writer.Write("\"" + kvp.Key + "\":[" + string.Join(",", kvp.Value) + "]");
            }
            writer.Write("\n  },\n");

            // GTIN Map
            writer.Write("  gtinMap: {\n");
            for (int i = 0; i < gtinMapLines.Count; i++) {
                if (i > 0) writer.Write(",\n");
                writer.Write(gtinMapLines[i]);
            }
            writer.Write("\n  },\n");

            // Helper function getConcept
            writer.Write("  getConcept: function(i) {\n");
            writer.Write("    var c = this.concepts[i];\n");
            writer.Write("    if (!c) return null;\n");
            writer.Write("    var nm = c[1];\n");
            writer.Write("    var vn = c[7];\n");
            writer.Write("    return {\n");
            writer.Write("      id: c[0],\n");
            writer.Write("      nm: nm,\n");
            writer.Write("      dg: c[2] || vn,\n");
            writer.Write("      st: c[3],\n");
            writer.Write("      fm: c[4],\n");
            writer.Write("      cm: c[5] ? [c[5]] : [],\n");
            writer.Write("      vt: c[6],\n");
            writer.Write("      vn: vn,\n");
            writer.Write("      vp: c[8],\n");
            writer.Write("      vpn: c[9] || nm,\n");
            writer.Write("      ap: c[10],\n");
            writer.Write("      apn: c[11],\n");
            writer.Write("      nti: Boolean(c[12]),\n");
            writer.Write("      cmm: Boolean(c[13])\n");
            writer.Write("    };\n");
            writer.Write("  },\n");

            // Helper function getGtin
            writer.Write("  getGtin: function(gtin) {\n");
            writer.Write("    if (!gtin) return null;\n");
            writer.Write("    var str = String(gtin).trim();\n");
            writer.Write("    var row = this.gtinMap[str];\n");
            writer.Write("    if (!row && str.length < 14) {\n");
            writer.Write("      row = this.gtinMap[str.padStart(14, '0')];\n");
            writer.Write("    }\n");
            writer.Write("    if (!row && str.length === 14 && str.charAt(0) === '0') {\n");
            writer.Write("      row = this.gtinMap[str.substring(1)];\n");
            writer.Write("    }\n");
            writer.Write("    if (!row) return null;\n");
            writer.Write("    var c = (row[5] >= 0) ? this.getConcept(row[5]) : null;\n");
            writer.Write("    var pSuffix = (row[2] >= 0 && row[2] < this.packs.length) ? this.packs[row[2]] : '';\n");
            writer.Write("    var pDesc = c ? ((c.apn || c.nm) + (pSuffix ? ' ' + pSuffix : '')) : pSuffix;\n");
            writer.Write("    return {\n");
            writer.Write("      amppId: row[0],\n");
            writer.Write("      ampId: row[1],\n");
            writer.Write("      vmpId: c ? c.vp : '',\n");
            writer.Write("      packDesc: pDesc,\n");
            writer.Write("      isHist: Boolean(row[3]),\n");
            writer.Write("      isDisc: Boolean(row[4]),\n");
            writer.Write("      conceptIndex: row[5], \n");
            writer.Write("      concept: c\n");
            writer.Write("    };\n");
            writer.Write("  }\n");
            writer.Write("};\n");
        }

        sw.Stop();
        var fi = new FileInfo(outJsPath);
        double sizeMb = Math.Round((double)fi.Length / (1024 * 1024), 2);

        var metricsJson = string.Format(
            "{{\n" +
            "  \"buildDurationMs\": {0},\n" +
            "  \"gtinTotal\": {1},\n" +
            "  \"gtinResolved\": {2},\n" +
            "  \"gtinResolutionPercent\": {3:F2},\n" +
            "  \"cmmTotalRows\": {4},\n" +
            "  \"cmmUniqueCodes\": {5},\n" +
            "  \"nationalVmpAdded\": {6},\n" +
            "  \"brandAmpAdded\": {7},\n" +
            "  \"totalConcepts\": {8},\n" +
            "  \"vtmSiblingGroups\": {9},\n" +
            "  \"uniquePackStrings\": {10},\n" +
            "  \"fileSizeBytes\": {11},\n" +
            "  \"fileSizeMb\": {12:F2},\n" +
            "  \"fileSizeUnder12Mb\": {13}\n" +
            "}}",
            sw.ElapsedMilliseconds, gtinTotal, gtinResolved,
            ((double)gtinResolved / gtinTotal) * 100.0,
            cmmLineCount, cmmCodeToConcept.Count, nationalVmpAdded, brandAmpAdded,
            concepts.Count, vtmSiblings.Count, packList.Count,
            fi.Length, sizeMb, (sizeMb < 12.0).ToString().ToLower()
        );
        File.WriteAllText(outMetricsPath, metricsJson, Encoding.UTF8);

        return string.Format("Built in {0}ms | Size: {1:F2} MB | GTINs: {2}/{3} resolved ({4:F1}%) | Concepts: {5}",
            sw.ElapsedMilliseconds, sizeMb, gtinResolved, gtinTotal,
            ((double)gtinResolved / gtinTotal) * 100.0, concepts.Count);
    }

    private static string Escape(string s) {
        if (string.IsNullOrEmpty(s)) return "";
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "").Replace("\n", " ");
    }

    private static string[] ParseCsvRow(string line) {
        var list = new List<string>();
        bool inQuotes = false;
        var cur = new StringBuilder();
        for (int i = 0; i < line.Length; i++) {
            char c = line[i];
            if (c == '\"') {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '\"') {
                    cur.Append('\"');
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (c == ',' && !inQuotes) {
                list.Add(cur.ToString().Trim());
                cur.Clear();
            } else {
                cur.Append(c);
            }
        }
        list.Add(cur.ToString().Trim());
        return list.ToArray();
    }
}
"@

Add-Type -TypeDefinition $csharpCode

$summary = [NearMissIndexCompiler]::Execute($DmdDir, $CmmFile, $OutJs, $OutMetrics)

Write-Host ""
Write-Host "Build Summary:" -ForegroundColor Green
Write-Host "  $summary" -ForegroundColor Green
Write-Host ""
Write-Host "Metrics written to : $OutMetrics"
Write-Host "Index written to   : $OutJs"
Write-Host "================================================================================" -ForegroundColor Cyan
