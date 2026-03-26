const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

// PANW Brand Colors
const NAVY = '141C26';
const ORANGE = 'FA582D';
const GRAY = '8A9BB0';
const RED = 'CC0000';
const GREEN = '00CC66';
const WHITE = 'FFFFFF';
const LIGHT_BG = 'F5F6F8';

// Load telemetry data
const dataPath = '/Users/johnshelest/Code/security-assessment-v3/last_parsed.json';
const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
const customer = "Reyes Holdings";
const shortName = "Reyes";
const dateStr = "March 2026";

let pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9';

// Define master slide for consistency (footer)
pptx.defineSlideMaster({
    title: 'MASTER_SLIDE',
    background: { color: WHITE },
    objects: [
        { text: { text: `PALO ALTO NETWORKS  ·  ${customer.toUpperCase()}  ·  ${dateStr.toUpperCase()}  ·  CONFIDENTIAL`, options: { x: 0.5, y: 7.0, w: 12, h: 0.3, color: GRAY, fontSize: 9 } } }
    ]
});

// --- SLIDE 1: Title ---
let slide1 = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
slide1.addText("PALO ALTO NETWORKS", { x: 0.5, y: 0.6, w: 8, h: 0.5, color: GRAY, fontSize: 12, bold: true });
slide1.addText(`${customer}\nSecurity Review`, { x: 0.5, y: 1.8, w: 10, h: 1.5, color: NAVY, fontSize: 44, bold: true });
slide1.addText(`${dateStr}  ·  Quarterly Business Review`, { x: 0.5, y: 3.3, w: 8, h: 0.5, color: ORANGE, fontSize: 18 });

let prepText = [
    { text: "Prepared by\n", options: { bold: true, color: GRAY } },
    { text: "John Shelest\nSolutions Consultant Palo Alto Networks\njshelest@paloaltonetworks.com" }
];
slide1.addText(prepText, { x: 0.5, y: 4.8, w: 4, h: 1, color: NAVY, fontSize: 12, lineSpacing: 18 });

let dataText = [
    { text: "Data sources\n", options: { bold: true, color: GRAY } },
    { text: `Panorama PAN-OS 11.1.13\n${(data.totalRows || 62956).toLocaleString()} threat log events\nAlienVault OTX enrichment\nSLR dataset  ·  2026-03-12` }
];
slide1.addText(dataText, { x: 4.5, y: 4.8, w: 5, h: 1, color: NAVY, fontSize: 12, lineSpacing: 18 });

// --- SLIDE 2: Attack Surface ---
let slide2 = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
slide2.addText("PALO ALTO NETWORKS", { x: 0.5, y: 0.4, w: 8, h: 0.3, color: GRAY, fontSize: 10, bold: true });
slide2.addText("ACT 1 OF 4  ·  ATTACK SURFACE", { x: 0.5, y: 0.7, w: 8, h: 0.4, color: ORANGE, fontSize: 14, bold: true });
slide2.addText("Your environment is more complex than peers.", { x: 0.5, y: 1.1, w: 10, h: 0.6, color: NAVY, fontSize: 28, bold: true });
slide2.addText("More applications. More access paths. More exposure. That complexity is what attackers look for.", { x: 0.5, y: 1.6, w: 12, h: 0.3, color: GRAY, fontSize: 14 });

let totalApps = data.slr?.totalApps || 739;
let remoteApps = data.slr?.remoteApps || 30;
let highRisk = data.slr?.highRiskApps || 58;
let saasApps = data.slr?.saasApps || 411;

let rows = [
    [
        { text: "", options: { fill: WHITE, border: {type:'none'} } },
        { text: customer.toUpperCase(), options: { fill: NAVY, color: WHITE, bold: true, align: 'center', valign: 'middle' } },
        { text: "INDUSTRY AVG  (T&L PEERS)", options: { fill: NAVY, color: WHITE, bold: true, align: 'center', valign: 'middle' } }
    ],
    [ {text: "APPLICATIONS", options: {bold: true}}, `${totalApps}`, "254" ],
    [ {text: "REMOTE ACCESS", options: {bold: true}}, `${remoteApps}`, "9" ],
    [ {text: "HIGH-RISK APPS", options: {bold: true}}, `${highRisk}`, "22" ],
    [ {text: "SAAS FOOTPRINT", options: {bold: true}}, `${saasApps}`, "134" ]
];

slide2.addTable(rows, { x: 0.5, y: 2.2, w: 9, h: 2.0, colW: [3, 3, 3], border: { pt: 1, color: 'E0E0E0' }, fontSize: 14, color: NAVY, align: 'center', valign: 'middle' });

let expText = [
    { text: `What this means for ${shortName}\n\n`, options: { bold: true, color: NAVY, fontSize: 16 } },
    { text: `${totalApps} applications means ${totalApps} policies to maintain.\n`, options: { bold: true, color: NAVY } },
    { text: `Every unmanaged app is a potential gap in your security policy — and at 3× the peer average, the odds of a misconfiguration increase significantly.\n\n` },
    { text: `${remoteApps} remote access tools is a governance problem, not just a security one.\n`, options: { bold: true, color: NAVY } },
    { text: `Consumer-grade tools bypass VPN and MFA controls. They're the #1 post-compromise persistence mechanism for ransomware operators.\n\n` },
    { text: `Complexity doesn't mean you're less secure — it means your controls need to work harder.\n`, options: { bold: true, color: NAVY } },
    { text: `The peer comparison tells us where to focus attention, not that something is broken.` }
];
slide2.addText(expText, { x: 0.5, y: 4.5, w: 12.33, h: 2.2, color: NAVY, fontSize: 12, fill: LIGHT_BG, align: 'left', valign: 'top', margin: 14 });

// Save
const outFile = path.join(__dirname, 'Reyes_CISO_QBR_Node.pptx');
pptx.writeFile({ fileName: outFile }).then(fileName => {
    console.log(`Successfully created: ${fileName}`);
});
