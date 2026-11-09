// MediaVault – COM682 Coursework 1 slide deck generator
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const SKILL = "/root/.claude/skills/synced/ddd4d19f-2459-4c46-8be9-c4efc9b55ed1_09076ea8-d13e-497f-b69e-b26895ce2aac/pptx";
const { applyTheme } = require(SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "MediaVault_COM682_CW1.pptx";

const THEME = {
  name: "MediaVault",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1C1E2E", lt1: "FFFFFF", dk2: "20244D", lt2: "F2F3F8",
    accent1: "FF6B5B", accent2: "15A396", accent3: "5B6CFF",
    accent4: "E8992A", accent5: "656A82", accent6: "D5D8E6",
    hlink: "3D4FE0", folHlink: "7A3DE0",
  },
};
const HEX = THEME.colors;

// ---------- icons ----------
const iconCache = {};
async function icon(name, hex) {
  const key = name + hex;
  if (iconCache[key]) return iconCache[key];
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(fa[name], { color: "#" + hex, size: 256 })
  );
  const buf = await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return (iconCache[key] = "image/png;base64," + buf.toString("base64"));
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "MediaVault – COM682 Coursework 1";
  pres.subject = "Cloud-native multimedia sharing platform design";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const S = pres.shapes;

  // ---------- layouts ----------
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.3, w: 12.1, h: 0.8, fontSize: 32, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
      { text: { text: "MediaVault · COM682 Cloud Native Development · Coursework 1", options: { x: 0.6, y: 7.05, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
    ],
    slideNumber: { x: 12.2, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
  });
  pres.defineSlideMaster({
    title: "DARK",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.3, w: 12.1, h: 0.8, fontSize: 32, bold: true, color: C.background1, valign: "middle", margin: 0 }, text: "" } },
    ],
  });

  // ---------- helpers ----------
  const T = (slide, text, o) => slide.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", fontSize: 14, color: C.text1 }, o));
  const shadow = () => ({ type: "outer", color: "20244D", blur: 6, offset: 2, angle: 90, opacity: 0.12 });
  function card(slide, x, y, w, h, o = {}) {
    slide.addShape(S.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.08, fill: { color: C.background1 }, line: { color: C.accent6, width: 0.75 }, shadow: shadow() }, o));
  }
  async function iconCircle(slide, name, x, y, d, colorKey, hexIcon = HEX.lt1) {
    slide.addShape(S.OVAL, { x, y, w: d, h: d, fill: { color: C[colorKey] }, line: { type: "none" } });
    const p = d * 0.24;
    slide.addImage({ data: await icon(name, hexIcon), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: name.replace(/^Fa/, "") + " icon" });
  }
  function arrow(slide, x1, y1, x2, y2, o = {}) {
    slide.addShape(S.LINE, {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
      flipH: x2 < x1, flipV: y2 < y1,
      line: { color: o.color || C.accent5, width: o.width || 1.5, dashType: o.dash || "solid", endArrowType: o.noHead ? undefined : "triangle", beginArrowType: o.both ? "triangle" : undefined },
    });
  }
  function badge(slide, n, x, y, colorKey = "accent1") {
    slide.addShape(S.OVAL, { x: x - 0.15, y: y - 0.15, w: 0.3, h: 0.3, fill: { color: C[colorKey] }, line: { color: C.background1, width: 1 } });
    T(slide, String(n), { x: x - 0.15, y: y - 0.15, w: 0.3, h: 0.3, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
  }
  async function node(slide, x, y, w, h, title, sub, iconName, colorKey) {
    card(slide, x, y, w, h);
    await iconCircle(slide, iconName, x + 0.12, y + (h - 0.5) / 2, 0.5, colorKey);
    T(slide, [
      { text: title, options: { bold: true, fontSize: 12, color: C.text2, breakLine: true } },
      { text: sub, options: { fontSize: 10, color: C.accent5 } },
    ], { x: x + 0.72, y: y + 0.08, w: w - 0.8, h: h - 0.16, valign: "middle" });
  }

  // =====================================================================
  // 1. TITLE
  // =====================================================================
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Introduction" });
  T(s, "COM682 CLOUD NATIVE DEVELOPMENT · COURSEWORK 1", { x: 0.8, y: 1.3, w: 7.5, h: 0.4, fontSize: 14, bold: true, color: C.accent1, charSpacing: 2 });
  T(s, "MediaVault", { x: 0.8, y: 1.9, w: 7.5, h: 1.3, fontSize: 66, bold: true, color: C.background1, fontFace: THEME.headFontFace });
  T(s, "A scalable, cloud-native platform for sharing photos and videos, designed on Microsoft Azure", { x: 0.8, y: 3.35, w: 6.9, h: 1.0, fontSize: 20, color: C.accent6 });
  T(s, [
    { text: "[Your Name]", options: { bold: true, fontSize: 18, color: C.background1, breakLine: true } },
    { text: "Student number: [B00XXXXXX]", options: { fontSize: 16, color: C.accent6 } },
  ], { x: 0.8, y: 5.0, w: 6.5, h: 0.9 });
  // visual cluster
  const tiles = [
    ["FaImage", "accent1", 8.7, 1.5], ["FaVideo", "accent2", 10.55, 1.5],
    ["FaCloud", "accent3", 8.7, 3.35], ["FaCloudUploadAlt", "accent4", 10.55, 3.35],
  ];
  for (const [ic, ck, x, y] of tiles) {
    s.addShape(S.ROUNDED_RECTANGLE, { x, y, w: 1.65, h: 1.65, rectRadius: 0.2, fill: { color: C[ck] }, line: { type: "none" } });
    s.addImage({ data: await icon(ic, HEX.lt1), x: x + 0.45, y: y + 0.45, w: 0.75, h: 0.75, altText: ic.replace(/^Fa/, "") + " icon" });
  }
  T(s, "Static web · Logic Apps · Functions · Blob · Cosmos DB · Azure SQL", { x: 8.7, y: 5.25, w: 3.5, h: 0.6, fontSize: 12, color: C.accent6 });

  // =====================================================================
  // 2. PROBLEM – traditional vs cloud
  // =====================================================================
  pres.addSection({ title: "Problem definition" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Problem definition" });
  s.addText("Why media sharing breaks on traditional hosting", { placeholder: "title" });
  T(s, "Users upload large, unpredictable volumes of images and video that are then read many times, worldwide. A single on-premises web server with local disks cannot absorb this.", { x: 0.6, y: 1.2, w: 12.1, h: 0.6, fontSize: 14, color: C.accent5 });
  T(s, "Traditional single-server hosting", { x: 0.6, y: 1.95, w: 5.4, h: 0.35, fontSize: 16, bold: true, color: C.accent1 });
  T(s, "Cloud-native response on Azure", { x: 7.3, y: 1.95, w: 5.4, h: 0.35, fontSize: 16, bold: true, color: C.accent2 });
  const probRows = [
    ["FaServer", "Fixed capacity", "Sized for average load; fails when a post goes viral, idle (but paid for) at night.", "FaBolt", "Elastic, event-driven compute", "Logic Apps and Functions add instances per request and scale back to zero [7], [8]."],
    ["FaHdd", "Files on local disk", "Storage fills up, is hard to grow and is a single point of failure.", "FaCloudUploadAlt", "Object storage", "Blob Storage holds massive unstructured data with built-in redundancy [5]."],
    ["FaGlobeEurope", "One location", "Users far from the server get slow page loads and video buffering.", "FaGlobe", "Global edge delivery", "Azure Front Door caches static files and thumbnails close to users [11]."],
    ["FaPoundSign", "Up-front CapEx", "Hardware bought ahead of demand; manual patching and backups.", "FaLayerGroup", "Managed, pay-as-you-go", "PaaS/serverless services: on-demand, measured service and rapid elasticity [1]."],
  ];
  for (let i = 0; i < probRows.length; i++) {
    const r = probRows[i]; const y = 2.4 + i * 1.12;
    card(s, 0.6, y, 5.6, 0.95, { fill: { color: C.accent1, transparency: 92 }, line: { type: "none" }, shadow: undefined });
    await iconCircle(s, r[0], 0.75, y + 0.2, 0.55, "accent1");
    T(s, [{ text: r[1], options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } }, { text: r[2], options: { fontSize: 12, color: C.text1 } }], { x: 1.45, y: y + 0.08, w: 4.65, h: 0.8, valign: "middle" });
    arrow(s, 6.35, y + 0.475, 7.15, y + 0.475, { color: C.accent5, width: 2 });
    card(s, 7.3, y, 5.4, 0.95, { fill: { color: C.accent2, transparency: 90 }, line: { type: "none" }, shadow: undefined });
    await iconCircle(s, r[3], 7.45, y + 0.2, 0.55, "accent2");
    T(s, [{ text: r[4], options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } }, { text: r[5], options: { fontSize: 12, color: C.text1 } }], { x: 8.15, y: y + 0.08, w: 4.45, h: 0.8, valign: "middle" });
  }

  // =====================================================================
  // 3. SCALABILITY PRESSURE POINTS
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Problem definition" });
  s.addText("Scalability issues for each resource", { placeholder: "title" });
  const pts = [
    ["FaBolt", "accent1", "Web/API tier: bursty traffic", "Viral posts multiply requests within minutes. Fixed VMs saturate CPU and connections; over-provisioning wastes money."],
    ["FaFileVideo", "accent3", "Upload path: large payloads", "Videos of hundreds of MB tie up API threads and bandwidth if proxied through the server."],
    ["FaHdd", "accent4", "Storage: unbounded growth", "Media is append-only and rarely deleted, so disk-based capacity planning always falls behind."],
    ["FaDatabase", "accent2", "Database: hot partitions", "Feeds, likes and comments are read-heavy and uneven; a single relational server becomes the bottleneck."],
    ["FaMicrochip", "accent1", "Processing: CPU spikes", "Thumbnails, tagging and moderation are CPU-heavy and must not block the user's request."],
    ["FaGlobe", "accent3", "Delivery: read-heavy, global", "Each upload is viewed many times; latency and egress grow with distance from one origin."],
  ];
  for (let i = 0; i < pts.length; i++) {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1, y = 1.35 + row * 2.35;
    card(s, x, y, 3.9, 2.1);
    await iconCircle(s, pts[i][0], x + 0.25, y + 0.25, 0.6, pts[i][1]);
    T(s, pts[i][2], { x: x + 1.0, y: y + 0.25, w: 2.75, h: 0.6, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    T(s, pts[i][3], { x: x + 0.25, y: y + 1.0, w: 3.45, h: 1.0, fontSize: 13, color: C.text1 });
  }
  s.addShape(S.ROUNDED_RECTANGLE, { x: 0.6, y: 6.15, w: 12.1, h: 0.65, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" } });
  T(s, [{ text: "Design goal: ", options: { bold: true, color: C.accent1 } }, { text: "stateless, event-driven tiers where every resource scales independently and only the cost of real usage is paid [4].", options: { color: C.background1 } }], { x: 0.85, y: 6.15, w: 11.7, h: 0.65, fontSize: 15, valign: "middle" });

  // =====================================================================
  // 4. SOLUTION ARCHITECTURE
  // =====================================================================
  pres.addSection({ title: "Solution design" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solution design" });
  s.addText("Solution architecture", { placeholder: "title" });
  const W = 1.95, H = 0.9;
  const cx = [0.45, 3.05, 5.6, 8.15, 10.75];
  const ry = [1.55, 2.85, 4.15];
  // data tier group
  s.addShape(S.ROUNDED_RECTANGLE, { x: 10.6, y: 1.4, w: 2.25, h: 3.97, rectRadius: 0.1, fill: { color: C.accent2, transparency: 92 }, line: { color: C.accent2, width: 1, dashType: "dash" } });
  T(s, "DATA TIER", { x: 10.6, y: 5.12, w: 2.25, h: 0.2, fontSize: 9, bold: true, color: C.accent2, align: "center" });
  // upload line (5)
  arrow(s, cx[0] + W / 2, ry[1], cx[0] + W / 2, 1.2, { noHead: true, color: C.accent1, width: 2 });
  arrow(s, cx[0] + W / 2, 1.2, cx[4] + W / 2, 1.2, { noHead: true, color: C.accent1, width: 2 });
  arrow(s, cx[4] + W / 2, 1.2, cx[4] + W / 2, ry[0], { color: C.accent1, width: 2 });
  badge(s, 5, 6.6, 1.2);
  T(s, "Direct upload with SAS (Valet Key) [3]", { x: 6.85, y: 1.24, w: 3.2, h: 0.3, fontSize: 10, bold: true, color: C.accent1 });
  // login line (1)
  arrow(s, cx[0] + W / 2, ry[1] + H, cx[0] + W / 2, ry[2] + H / 2, { noHead: true });
  arrow(s, cx[0] + W / 2, ry[2] + H / 2, cx[2], ry[2] + H / 2);
  badge(s, 1, 3.0, ry[2] + H / 2);
  // 2, 3: browser -> FD -> static / logic apps
  arrow(s, cx[0] + W, ry[1] + H / 2, cx[1], ry[1] + H / 2);
  badge(s, 2, (cx[0] + W + cx[1]) / 2, ry[1] + H / 2 - 0.22);
  arrow(s, cx[1] + W, ry[1] + 0.25, cx[2], ry[0] + H / 2);
  arrow(s, cx[1] + W, ry[1] + H / 2, cx[2], ry[1] + H / 2);
  badge(s, 3, (cx[1] + W + cx[2]) / 2, ry[1] + H / 2 + 0.22);
  // 4: logic apps -> data
  arrow(s, cx[2] + W, ry[1] + H / 2, 10.6, ry[1] + H / 2, { color: C.accent2, width: 2 });
  badge(s, 4, 9.1, ry[1] + H / 2, "accent2");
  // 6/8: blob <-> functions
  arrow(s, cx[4], ry[0] + H / 2, cx[3] + W, ry[0] + H / 2, { both: true, color: C.accent3, width: 2 });
  badge(s, 6, (cx[3] + W + cx[4]) / 2, ry[0] + H / 2 - 0.25, "accent3");
  badge(s, 8, (cx[3] + W + cx[4]) / 2, ry[0] + H / 2 + 0.25, "accent3");
  // 7: functions -> AI vision
  arrow(s, cx[3] + W / 2 + 0.45, ry[0] + H, cx[3] + W / 2 + 0.45, ry[2], { color: C.accent3, width: 2, both: true });
  badge(s, 7, cx[3] + W / 2 + 0.45, ry[2] - 0.35, "accent3");
  // nodes
  await node(s, cx[0], ry[1], W, H, "User browser", "Desktop / mobile", "FaLaptop", "text2");
  await node(s, cx[1], ry[1], W, H, "Azure Front Door", "CDN cache + WAF [11]", "FaGlobe", "accent3");
  await node(s, cx[2], ry[0], W, H, "Static website", "HTML/JS in Blob $web [6]", "FaHtml5", "accent4");
  await node(s, cx[2], ry[1], W, H, "Logic Apps", "REST CRUD API [7]", "FaSitemap", "accent2");
  await node(s, cx[2], ry[2], W, H, "Entra External ID", "Sign-in, OAuth tokens [12]", "FaUserLock", "text2");
  await node(s, cx[3], ry[0], W, H, "Azure Functions", "Flex Consumption [8]", "FaBolt", "accent3");
  await node(s, cx[3], ry[2], W, H, "Azure AI Vision", "Tags + moderation [13]", "FaRobot", "accent3");
  await node(s, cx[4], ry[0], W, H, "Blob Storage", "originals, thumbnails [5]", "FaCloudUploadAlt", "accent1");
  await node(s, cx[4], ry[1], W, H, "Cosmos DB", "Media metadata (NoSQL)", "FaDatabase", "accent2");
  await node(s, cx[4], ry[2], W, H, "Azure SQL Database", "Users, albums (SQL)", "FaDatabase", "accent4");
  // cross-cutting
  T(s, "CROSS-CUTTING", { x: 0.45, y: 5.42, w: 1.8, h: 0.35, fontSize: 10, bold: true, color: C.accent5, valign: "middle" });
  const chips = [["FaChartLine", "Application Insights + Azure Monitor"], ["FaKey", "Key Vault + managed identities"], ["FaGithub", "GitHub Actions CI/CD"]];
  for (let i = 0; i < chips.length; i++) {
    const x = 2.3 + i * 3.55;
    s.addShape(S.ROUNDED_RECTANGLE, { x, y: 5.4, w: 3.35, h: 0.4, rectRadius: 0.2, fill: { color: C.background2 }, line: { color: C.accent6, width: 0.75 } });
    s.addImage({ data: await icon(chips[i][0], HEX.dk2), x: x + 0.12, y: 5.47, w: 0.26, h: 0.26, altText: "icon" });
    T(s, chips[i][1], { x: x + 0.48, y: 5.4, w: 2.8, h: 0.4, fontSize: 11, color: C.text2, valign: "middle" });
  }
  // legend
  const legend = [
    "1  Sign in, receive OAuth token", "2  Load app via Front Door (cached)", "3  REST calls with bearer token", "4  CRUD on Cosmos DB, SQL, Blob (SAS)",
    "5  Browser uploads file straight to Blob", "6  Blob-created event triggers Function", "7  Function calls AI Vision (tags, safety)", "8  Thumbnail → Blob, tags → Cosmos DB",
  ];
  for (let i = 0; i < legend.length; i++) {
    const col = i % 4, row = Math.floor(i / 4);
    T(s, legend[i], { x: 0.45 + col * 3.1, y: 6.0 + row * 0.42, w: 3.0, h: 0.35, fontSize: 11, color: C.text1, valign: "middle" });
  }

  // =====================================================================
  // 5. NETWORK / SECURITY VIEW
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solution design" });
  s.addText("Network diagram and security boundaries", { placeholder: "title" });
  // zones
  const zone = (x, y, w, h, label, ck, fillT = 94) => {
    s.addShape(S.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C[ck], transparency: fillT }, line: { color: C[ck], width: 1, dashType: "dash" } });
    T(s, label, { x: x + 0.12, y: y + 0.08, w: w - 0.24, h: 0.3, fontSize: 10, bold: true, color: C[ck] });
  };
  zone(0.45, 1.45, 1.95, 5.0, "INTERNET", "accent5");
  zone(2.65, 1.45, 2.15, 5.0, "AZURE EDGE (GLOBAL)", "accent3");
  zone(5.05, 1.45, 7.8, 5.0, "AZURE REGION: UK SOUTH · rg-mediavault-prod", "text2", 97);
  zone(5.25, 1.9, 3.6, 2.0, "WEB + API TIER (PaaS)", "accent4");
  zone(5.25, 4.05, 3.6, 1.45, "VNet 10.0.0.0/16 · snet-func 10.0.1.0/24", "accent3");
  zone(9.1, 1.9, 3.55, 3.6, "DATA TIER (firewall: selected networks)", "accent2");
  await node(s, 0.55, 3.35, 1.75, 0.9, "Users", "HTTPS only", "FaUser", "text2");
  await node(s, 2.75, 3.35, 1.95, 0.9, "Front Door", "WAF · TLS 1.2+ to origins", "FaShieldAlt", "accent3");
  await node(s, 5.4, 2.3, 3.3, 0.7, "Static website ($web)", "Origin for HTML/JS/CSS", "FaHtml5", "accent4");
  await node(s, 5.4, 3.1, 3.3, 0.7, "Logic Apps (Consumption)", "Inbound IPs: Front Door only", "FaSitemap", "accent2");
  await node(s, 5.4, 4.45, 3.3, 0.9, "Functions (Flex)", "VNet-integrated outbound", "FaBolt", "accent3");
  await node(s, 9.25, 2.3, 3.25, 0.8, "Blob Storage", "No anonymous access; SAS only", "FaCloudUploadAlt", "accent1");
  await node(s, 9.25, 3.3, 3.25, 0.8, "Cosmos DB", "Entra RBAC, HTTPS 443", "FaDatabase", "accent2");
  await node(s, 9.25, 4.3, 3.25, 0.8, "Azure SQL", "Entra auth, TDS 1433", "FaDatabase", "accent4");
  // arrows with protocols
  arrow(s, 2.3, 3.8, 2.75, 3.8, { width: 2 });
  arrow(s, 4.7, 3.6, 5.4, 2.65, { width: 2 });
  arrow(s, 4.7, 3.8, 5.4, 3.45, { width: 2 });
  arrow(s, 8.7, 3.45, 9.25, 3.7, { color: C.accent2, width: 2 });
  arrow(s, 8.7, 4.9, 9.25, 4.7, { color: C.accent3, width: 2 });
  arrow(s, 8.7, 4.7, 9.25, 2.9, { color: C.accent3, width: 2 });
  // SAS route along bottom
  arrow(s, 1.425, 4.25, 1.425, 5.8, { noHead: true, color: C.accent1, width: 2 });
  arrow(s, 1.425, 5.8, 12.75, 5.8, { noHead: true, color: C.accent1, width: 2 });
  arrow(s, 12.75, 5.8, 12.75, 2.7, { noHead: true, color: C.accent1, width: 2 });
  arrow(s, 12.75, 2.7, 12.5, 2.7, { color: C.accent1, width: 2 });
  T(s, "Direct upload: HTTPS 443 + short-lived, write-only SAS token", { x: 2.7, y: 5.85, w: 6.0, h: 0.3, fontSize: 10, bold: true, color: C.accent1 });
  // security strip
  const sec = [["FaUserLock", "Entra External ID: tokens on every API call"], ["FaKey", "Key Vault + managed identities: no secrets in code"], ["FaChartLine", "Log Analytics + App Insights: all traffic logged"]];
  for (let i = 0; i < sec.length; i++) {
    const x = 0.45 + i * 4.18;
    s.addShape(S.ROUNDED_RECTANGLE, { x, y: 6.55, w: 4.0, h: 0.4, rectRadius: 0.2, fill: { color: C.background2 }, line: { color: C.accent6, width: 0.75 } });
    s.addImage({ data: await icon(sec[i][0], HEX.dk2), x: x + 0.12, y: 6.62, w: 0.26, h: 0.26, altText: "icon" });
    T(s, sec[i][1], { x: x + 0.48, y: 6.55, w: 3.45, h: 0.4, fontSize: 11, color: C.text2, valign: "middle" });
  }

  // =====================================================================
  // 6. WIREFRAMES
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solution design" });
  s.addText("Website layout wireframes", { placeholder: "title" });
  const G = C.accent6, GF = C.background2;
  const box = (x, y, w, h, o = {}) => s.addShape(o.round ? S.ROUNDED_RECTANGLE : S.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: GF }, line: { color: G, width: 0.75 } }, o.round ? { rectRadius: 0.05 } : {}, o.extra || {}));
  const lbl = (t, x, y, w, h, o = {}) => T(s, t, Object.assign({ x, y, w, h, fontSize: 9, color: C.accent5, valign: "middle" }, o));
  const frames = [0.6, 4.7, 8.8];
  const FW = 3.9, FY = 1.3, FH = 4.75;
  for (const fx of frames) {
    box(fx, FY, FW, FH, { extra: { fill: { color: C.background1 }, line: { color: C.accent5, width: 1 } } });
    box(fx, FY, FW, 0.25, { extra: { fill: { color: C.accent6 }, line: { type: "none" } } });
    for (let d = 0; d < 3; d++) s.addShape(S.OVAL, { x: fx + 0.1 + d * 0.15, y: FY + 0.08, w: 0.09, h: 0.09, fill: { color: C.accent5 }, line: { type: "none" } });
    // nav
    lbl("MediaVault", fx + 0.15, FY + 0.33, 1.0, 0.3, { bold: true, color: C.text2, fontSize: 10 });
    box(fx + 1.2, FY + 0.36, 1.35, 0.24, { round: true, extra: { fill: { color: C.background1 } } });
    lbl("Search…", fx + 1.28, FY + 0.36, 1.2, 0.24, { fontSize: 8 });
    s.addShape(S.ROUNDED_RECTANGLE, { x: fx + 2.65, y: FY + 0.36, w: 0.7, h: 0.24, rectRadius: 0.05, fill: { color: C.accent1 }, line: { type: "none" } });
    lbl("Upload", fx + 2.65, FY + 0.36, 0.7, 0.24, { fontSize: 8, color: C.background1, bold: true, align: "center" });
    s.addShape(S.OVAL, { x: fx + 3.5, y: FY + 0.35, w: 0.26, h: 0.26, fill: { color: G }, line: { type: "none" } });
    s.addShape(S.LINE, { x: fx, y: FY + 0.7, w: FW, h: 0, line: { color: G, width: 0.75 } });
  }
  // frame 1: feed
  let fx = frames[0];
  ["All", "Images", "Videos"].forEach((t, i) => {
    s.addShape(S.ROUNDED_RECTANGLE, { x: fx + 0.15 + i * 0.75, y: FY + 0.85, w: 0.65, h: 0.24, rectRadius: 0.1, fill: { color: i === 0 ? C.text2 : C.background2 }, line: { color: G, width: 0.5 } });
    lbl(t, fx + 0.15 + i * 0.75, FY + 0.85, 0.65, 0.24, { align: "center", fontSize: 8, color: i === 0 ? C.background1 : C.accent5 });
  });
  for (let r = 0; r < 3; r++) for (let c = 0; c < 3; c++) {
    const x = fx + 0.15 + c * 1.22, y = FY + 1.25 + r * 1.15;
    box(x, y, 1.12, 0.8);
    s.addShape(S.LINE, { x, y, w: 1.12, h: 0.8, line: { color: G, width: 0.5 } });
    s.addShape(S.LINE, { x, y, w: 1.12, h: 0.8, flipV: true, line: { color: G, width: 0.5 } });
    if ((r + c) % 2 === 1) s.addShape(S.OVAL, { x: x + 0.43, y: y + 0.27, w: 0.26, h: 0.26, fill: { color: C.accent1 }, line: { type: "none" } });
    box(x, y + 0.86, 0.8, 0.08, { extra: { fill: { color: G }, line: { type: "none" } } });
  }
  // frame 2: upload
  fx = frames[1];
  box(fx + 0.2, FY + 0.9, FW - 0.4, 1.3, { round: true, extra: { fill: { color: C.accent1, transparency: 92 }, line: { color: C.accent1, width: 1, dashType: "dash" } } });
  s.addImage({ data: await icon("FaCloudUploadAlt", HEX.accent1), x: fx + FW / 2 - 0.22, y: FY + 1.0, w: 0.44, h: 0.44, altText: "Upload icon" });
  lbl("Drag & drop an image or video (max 500 MB)", fx + 0.3, FY + 1.5, FW - 0.6, 0.3, { align: "center", fontSize: 9 });
  lbl("or browse files", fx + 0.3, FY + 1.78, FW - 0.6, 0.25, { align: "center", fontSize: 8, color: C.accent1 });
  const fields = ["Title", "Description", "Album  ▾", "Visibility: Public / Private"];
  fields.forEach((f, i) => {
    lbl(f, fx + 0.2, FY + 2.35 + i * 0.5, 2.5, 0.2, { fontSize: 8 });
    box(fx + 0.2, FY + 2.55 + i * 0.5, FW - 0.4, 0.24, { round: true, extra: { fill: { color: C.background1 } } });
  });
  box(fx + 0.2, FY + 4.4, FW - 0.4, 0.1, { round: true, extra: { fill: { color: G }, line: { type: "none" } } });
  box(fx + 0.2, FY + 4.4, 2.0, 0.1, { round: true, extra: { fill: { color: C.accent2 }, line: { type: "none" } } });
  s.addShape(S.ROUNDED_RECTANGLE, { x: fx + FW - 1.1, y: FY + 4.15, w: 0.9, h: 0.2, rectRadius: 0.05, fill: { color: C.accent2 }, line: { type: "none" } });
  lbl("Publish", fx + FW - 1.1, FY + 4.15, 0.9, 0.2, { fontSize: 8, color: C.background1, bold: true, align: "center" });
  lbl("Uploading 62%", fx + 0.2, FY + 4.12, 1.5, 0.24, { fontSize: 8 });
  // frame 3: detail
  fx = frames[2];
  box(fx + 0.2, FY + 0.85, FW - 0.4, 1.75);
  s.addShape(S.OVAL, { x: fx + FW / 2 - 0.25, y: FY + 1.48, w: 0.5, h: 0.5, fill: { color: C.accent1 }, line: { type: "none" } });
  s.addImage({ data: await icon("FaPlay", HEX.lt1), x: fx + FW / 2 - 0.08, y: FY + 1.61, w: 0.22, h: 0.24, altText: "Play icon" });
  lbl("Giant's Causeway at sunset", fx + 0.2, FY + 2.68, 2.6, 0.25, { fontSize: 10, bold: true, color: C.text2 });
  ["coast", "sunset", "rocks"].forEach((t, i) => {
    s.addShape(S.ROUNDED_RECTANGLE, { x: fx + 0.2 + i * 0.68, y: FY + 2.98, w: 0.6, h: 0.2, rectRadius: 0.08, fill: { color: C.accent3, transparency: 80 }, line: { type: "none" } });
    lbl("#" + t, fx + 0.2 + i * 0.68, FY + 2.98, 0.6, 0.2, { fontSize: 7, align: "center", color: C.accent3 });
  });
  lbl("AI tags", fx + 2.3, FY + 2.98, 0.8, 0.2, { fontSize: 7, italic: true });
  lbl("♥ 128   💬 14", fx + 0.2, FY + 3.25, 1.6, 0.22, { fontSize: 8 });
  [["Edit", C.accent4], ["Delete", C.accent1]].forEach(([t, col], i) => {
    s.addShape(S.ROUNDED_RECTANGLE, { x: fx + 2.1 + i * 0.8, y: FY + 3.25, w: 0.7, h: 0.22, rectRadius: 0.05, fill: { color: col }, line: { type: "none" } });
    lbl(t, fx + 2.1 + i * 0.8, FY + 3.25, 0.7, 0.22, { fontSize: 8, color: C.background1, bold: true, align: "center" });
  });
  for (let i = 0; i < 3; i++) {
    s.addShape(S.OVAL, { x: fx + 0.2, y: FY + 3.62 + i * 0.36, w: 0.24, h: 0.24, fill: { color: G }, line: { type: "none" } });
    box(fx + 0.55, FY + 3.66 + i * 0.36, 2.6 - i * 0.5, 0.14, { extra: { fill: { color: G }, line: { type: "none" } } });
  }
  // captions
  const caps = [["1  Home feed", "GET /api/media?type=&tag="], ["2  Upload", "POST /api/media → SAS upload to Blob"], ["3  Media detail (owner view)", "GET · PUT · DELETE /api/media/{id}"]];
  caps.forEach((c, i) => T(s, [{ text: c[0], options: { bold: true, fontSize: 13, color: C.text2, breakLine: true } }, { text: c[1], options: { fontSize: 11, color: C.accent5, fontFace: "Courier New" } }], { x: frames[i], y: 6.15, w: FW, h: 0.7 }));

  // =====================================================================
  // 7. DATABASE RESOURCES + SQL ERD
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solution design" });
  s.addText("Database design: resources and relational ERD", { placeholder: "title" });
  card(s, 0.6, 1.35, 4.6, 5.45, { fill: { color: C.background2 }, line: { type: "none" }, shadow: undefined });
  T(s, "Data resources", { x: 0.85, y: 1.5, w: 4.1, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  const res = [
    ["Azure SQL Database", "sqldb-mediavault (serverless, General Purpose) [10]", "Tables: Users, Albums, Follows", "accent4"],
    ["Azure Cosmos DB for NoSQL", "cosmos-mediavault › database mediadb [9]", "Containers: media (/userId), comments (/mediaId)", "accent2"],
    ["Azure Blob Storage", "stmediavault (StorageV2, Hot tier) [5]", "Containers: originals (private), thumbnails, $web", "accent1"],
    ["Indexes & integrity", "Unique index on EntraObjectId and Email; index on OwnerId", "ON DELETE CASCADE for a user's albums and follows", "accent3"],
  ];
  res.forEach((r, i) => {
    const y = 2.05 + i * 1.18;
    s.addShape(S.OVAL, { x: 0.85, y: y + 0.08, w: 0.2, h: 0.2, fill: { color: C[r[3]] }, line: { type: "none" } });
    T(s, [
      { text: r[0], options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } },
      { text: r[1], options: { fontSize: 12, color: C.text1, breakLine: true } },
      { text: r[2], options: { fontSize: 12, color: C.accent5 } },
    ], { x: 1.2, y, w: 3.85, h: 1.35 });
  });
  // ERD
  function entity(x, y, w, name, rows, ck) {
    const rh = 0.3, h = 0.42 + rows.length * rh;
    s.addShape(S.RECTANGLE, { x, y, w, h, fill: { color: C.background1 }, line: { color: C[ck], width: 1.25 }, shadow: shadow() });
    s.addShape(S.RECTANGLE, { x, y, w, h: 0.42, fill: { color: C[ck] }, line: { color: C[ck], width: 1.25 } });
    T(s, name, { x: x + 0.12, y, w: w - 0.24, h: 0.42, fontSize: 14, bold: true, color: C.background1, valign: "middle" });
    rows.forEach((r, i) => {
      T(s, [
        { text: r[0], options: { bold: true, color: C.text2 } },
        { text: "  " + r[1], options: { color: C.accent5 } },
      ], { x: x + 0.12, y: y + 0.45 + i * rh, w: w - 0.9, h: rh, fontSize: 11, valign: "middle" });
      if (r[2]) T(s, r[2], { x: x + w - 0.75, y: y + 0.45 + i * rh, w: 0.65, h: rh, fontSize: 10, bold: true, color: C[ck], align: "right", valign: "middle" });
    });
    return h;
  }
  entity(5.6, 1.45, 3.4, "Users", [["UserId", "uniqueidentifier", "PK"], ["EntraObjectId", "nvarchar(64)", "UQ"], ["DisplayName", "nvarchar(50)", ""], ["Email", "nvarchar(256)", "UQ"], ["CreatedAt", "datetime2", ""]], "accent4");
  entity(9.6, 1.45, 3.1, "Albums", [["AlbumId", "int identity", "PK"], ["OwnerId", "uniqueidentifier", "FK"], ["Title", "nvarchar(100)", ""], ["Visibility", "varchar(10)", ""], ["CreatedAt", "datetime2", ""]], "accent4");
  entity(5.6, 4.35, 3.4, "Follows", [["FollowerId", "uniqueidentifier", "PK, FK"], ["FolloweeId", "uniqueidentifier", "PK, FK"], ["CreatedAt", "datetime2", ""]], "accent4");
  // relationships
  arrow(s, 9.0, 2.2, 9.6, 2.2, { noHead: true, width: 1.5, color: C.text2 });
  T(s, "1", { x: 9.03, y: 1.9, w: 0.2, h: 0.25, fontSize: 12, bold: true, color: C.text2 });
  T(s, "*", { x: 9.4, y: 1.9, w: 0.2, h: 0.25, fontSize: 14, bold: true, color: C.text2 });
  T(s, "owns", { x: 9.0, y: 2.27, w: 0.6, h: 0.25, fontSize: 10, italic: true, color: C.accent5, align: "center" });
  arrow(s, 6.6, 3.37, 6.6, 4.35, { noHead: true, width: 1.5, color: C.text2 });
  arrow(s, 8.0, 3.37, 8.0, 4.35, { noHead: true, width: 1.5, color: C.text2 });
  T(s, "1", { x: 6.68, y: 3.4, w: 0.2, h: 0.25, fontSize: 12, bold: true, color: C.text2 });
  T(s, "*", { x: 6.68, y: 4.07, w: 0.2, h: 0.25, fontSize: 14, bold: true, color: C.text2 });
  T(s, "1", { x: 8.08, y: 3.4, w: 0.2, h: 0.25, fontSize: 12, bold: true, color: C.text2 });
  T(s, "*", { x: 8.08, y: 4.07, w: 0.2, h: 0.25, fontSize: 14, bold: true, color: C.text2 });
  T(s, "follower", { x: 5.75, y: 3.72, w: 0.8, h: 0.25, fontSize: 10, italic: true, color: C.accent5, align: "right" });
  T(s, "followee", { x: 8.1, y: 3.72, w: 0.8, h: 0.25, fontSize: 10, italic: true, color: C.accent5 });
  card(s, 9.6, 4.35, 3.1, 2.45, { fill: { color: C.accent4, transparency: 88 }, line: { type: "none" }, shadow: undefined });
  T(s, [
    { text: "Why relational here?", options: { bold: true, fontSize: 13, color: C.text2, breakLine: true } },
    { text: "Accounts and the follow graph need unique constraints, foreign keys and joins. Media IDs in Cosmos DB reference UserId and AlbumId.", options: { fontSize: 12, color: C.text1 } },
  ], { x: 9.8, y: 4.5, w: 2.75, h: 2.2 });

  // =====================================================================
  // 8. NOSQL JSON
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solution design" });
  s.addText("NoSQL design: Cosmos DB documents", { placeholder: "title" });
  let codeFs = 13;
  const codeRuns = (lines) => {
    const codeOpts = { fontFace: "Courier New", fontSize: codeFs, color: C.background2 };
    const cm = { fontFace: "Courier New", fontSize: codeFs, color: C.accent2 };
    const runs = [];
    lines.forEach((l, i) => {
      const last = i === lines.length - 1;
      if (l[1]) { runs.push({ text: l[0], options: Object.assign({}, codeOpts) }); runs.push({ text: l[1], options: Object.assign({ breakLine: !last }, cm) }); }
      else runs.push({ text: l[0], options: Object.assign({ breakLine: !last }, codeOpts) });
    });
    return runs;
  };
  s.addShape(S.ROUNDED_RECTANGLE, { x: 0.6, y: 1.35, w: 6.6, h: 5.45, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" } });
  T(s, "container: media   ·   partition key: /userId", { x: 0.85, y: 1.48, w: 6.2, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
  T(s, codeRuns([
    ["{"],
    ['  "id": "m_7f3a9c",', "         // unique media id"],
    ['  "userId": "u_102",', "        // partition key"],
    ['  "albumId": 14,', "            // → SQL Albums"],
    ['  "type": "video",', "          // image | video"],
    ['  "title": "Giant\'s Causeway",'],
    ['  "blobPath": "originals/u_102/m_7f3a9c.mp4",'],
    ['  "thumbPath": "thumbnails/u_102/m_7f3a9c.jpg",'],
    ['  "contentType": "video/mp4",'],
    ['  "sizeBytes": 48211337,'],
    ['  "durationSec": 42,', "        // null for images"],
    ['  "tags": ["coast", "sunset"],', " // AI Vision"],
    ['  "moderation": "approved",', "  // pending|rejected"],
    ['  "likeCount": 128,'],
    ['  "createdAt": "2026-11-09T18:22:04Z"'],
    ["}"],
  ]), { x: 0.85, y: 1.85, w: 6.25, h: 4.85, lineSpacingMultiple: 1.0 });
  codeFs = 11;
  s.addShape(S.ROUNDED_RECTANGLE, { x: 7.5, y: 1.35, w: 5.2, h: 2.65, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" } });
  T(s, "container: comments  ·  partition key: /mediaId", { x: 7.7, y: 1.48, w: 4.9, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
  T(s, codeRuns([
    ["{"],
    ['  "id": "c_91be",'],
    ['  "mediaId": "m_7f3a9c",', " // partition key"],
    ['  "userId": "u_205",'],
    ['  "text": "Stunning shot!",'],
    ['  "createdAt": "2026-11-09T19:01:47Z"'],
    ["}"],
  ]), { x: 7.7, y: 1.85, w: 4.95, h: 2.05 });
  card(s, 7.5, 4.2, 5.2, 2.6, { fill: { color: C.accent2, transparency: 90 }, line: { type: "none" }, shadow: undefined });
  T(s, [
    { text: "Why NoSQL for metadata?", options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } },
    { text: "Flexible schema: video-only fields without empty columns", options: { bullet: true, fontSize: 13, breakLine: true } },
    { text: "/userId spreads writes across partitions; a user's gallery is a single-partition query [9]", options: { bullet: true, fontSize: 13, breakLine: true } },
    { text: "Comments partitioned by /mediaId so a post's thread loads in one read", options: { bullet: true, fontSize: 13, breakLine: true } },
    { text: "Autoscale RU/s absorbs viral spikes", options: { bullet: true, fontSize: 13 } },
  ], { x: 7.7, y: 4.32, w: 4.85, h: 2.4, paraSpaceAfter: 3 });

  // =====================================================================
  // 9. REST API UML
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solution design" });
  s.addText("REST API design (UML class diagram)", { placeholder: "title" });
  function umlClass(x, y, w, stereo, name, attrs, ops, ck, fs = 11) {
    const lh = fs === 11 ? 0.24 : 0.22;
    const hh = 0.6, ah = attrs.length ? attrs.length * lh + 0.12 : 0.15, oh = ops.length * lh + 0.12;
    const h = hh + ah + oh;
    s.addShape(S.RECTANGLE, { x, y, w, h, fill: { color: C.background1 }, line: { color: C[ck], width: 1.25 }, shadow: shadow() });
    s.addShape(S.RECTANGLE, { x, y, w, h: hh, fill: { color: C[ck], transparency: 85 }, line: { color: C[ck], width: 1.25 } });
    T(s, [{ text: stereo, options: { fontSize: 10, italic: true, color: C.accent5, breakLine: true } }, { text: name, options: { fontSize: 14, bold: true, color: C.text2 } }], { x, y: y + 0.04, w, h: hh - 0.08, align: "center", valign: "middle" });
    s.addShape(S.LINE, { x, y: y + hh + ah, w, h: 0, line: { color: C[ck], width: 1 } });
    if (attrs.length) T(s, attrs.join("\n"), { x: x + 0.12, y: y + hh + 0.06, w: w - 0.24, h: ah - 0.08, fontSize: fs, fontFace: "Courier New", color: C.text1 });
    T(s, ops.join("\n"), { x: x + 0.12, y: y + hh + ah + 0.06, w: w - 0.24, h: oh - 0.08, fontSize: fs, fontFace: "Courier New", color: C.text1 });
    return h;
  }
  umlClass(0.6, 1.35, 6.3, "«Logic App · HTTP trigger»", "MediaApi", ["base: https://mediavault.azurefd.net/api"], [
    "+ POST   /media             : 201 {id, uploadUrl}",
    "+ GET    /media?type&tag    : 200 Media[]",
    "+ GET    /media/{id}        : 200 Media",
    "+ PUT    /media/{id}        : 200 Media",
    "+ DELETE /media/{id}        : 204",
    "+ GET    /media/{id}/comments: 200 Comment[]",
    "+ POST   /media/{id}/comments: 201 Comment",
  ], "accent2", 11);
  umlClass(0.6, 4.55, 6.3, "«Logic App · HTTP trigger»", "UserApi", [], [
    "+ GET    /users/{id}        : 200 User",
    "+ PUT    /users/me          : 200 User",
    "+ POST   /users/{id}/follow : 201",
    "+ DELETE /users/{id}/follow : 204",
  ], "accent2", 11);
  umlClass(7.75, 1.35, 2.35, "«Cosmos DB»", "Media", ["id: string", "userId: string", "type: MediaType", "title: string", "blobPath: string", "tags: string[]"], [], "accent1", 10);
  umlClass(10.45, 1.35, 2.3, "«Cosmos DB»", "Comment", ["id: string", "mediaId: string", "userId: string", "text: string"], [], "accent1", 10);
  umlClass(7.75, 4.55, 2.35, "«Azure SQL»", "User", ["UserId: guid", "DisplayName", "Email"], [], "accent4", 10);
  // associations
  arrow(s, 6.9, 2.4, 7.75, 2.4, { dash: "dash", color: C.accent5 });
  T(s, "«uses»", { x: 6.9, y: 2.1, w: 0.85, h: 0.25, fontSize: 9, italic: true, color: C.accent5, align: "center" });
  arrow(s, 6.9, 5.3, 7.75, 5.3, { dash: "dash", color: C.accent5 });
  T(s, "«uses»", { x: 6.9, y: 5.0, w: 0.85, h: 0.25, fontSize: 9, italic: true, color: C.accent5, align: "center" });
  arrow(s, 10.1, 2.3, 10.45, 2.3, { noHead: true, color: C.text2 });
  T(s, "1", { x: 10.1, y: 2.0, w: 0.15, h: 0.25, fontSize: 11, bold: true, color: C.text2 });
  T(s, "*", { x: 10.32, y: 2.0, w: 0.15, h: 0.25, fontSize: 13, bold: true, color: C.text2 });
  arrow(s, 8.9, 4.55, 8.9, 3.53, { noHead: true, color: C.text2 });
  T(s, "1", { x: 9.0, y: 4.27, w: 0.2, h: 0.25, fontSize: 11, bold: true, color: C.text2 });
  T(s, "*", { x: 9.0, y: 3.55, w: 0.2, h: 0.25, fontSize: 13, bold: true, color: C.text2 });
  card(s, 10.45, 3.7, 2.3, 2.45, { fill: { color: C.accent2, transparency: 90 }, line: { type: "none" }, shadow: undefined });
  T(s, [
    { text: "Each operation", options: { bold: true, fontSize: 12, color: C.text2, breakLine: true } },
    { text: "= one workflow: HTTP Request trigger → Cosmos / SQL / Blob connector → Response [7]. Calls require an Entra bearer token (Logic Apps authorization policy).", options: { fontSize: 11, color: C.text1 } },
  ], { x: 10.6, y: 3.82, w: 2.05, h: 2.25 });

  // =====================================================================
  // 10. TECHNOLOGY CHOICES
  // =====================================================================
  pres.addSection({ title: "Critical evaluation" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Critical evaluation" });
  s.addText("Why these services? Alternatives considered", { placeholder: "title" });
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 12 } });
  const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontSize: 11, color: C.text1 }, o) });
  const rowsT = [
    ["Front-end hosting", "Blob static website + Front Door", "App Service web app", "No servers to run, global caching, very low cost [6]", "No server-side rendering"],
    ["REST CRUD API", "Logic Apps (Consumption)", "App Service API; Functions HTTP", "Visual workflows, built-in connectors, pay per action [7]", "Higher latency, harder to unit-test"],
    ["Media processing", "Azure Functions (Flex)", "VM / App Service WebJobs", "Event-driven, scales per event, scales to zero [8]", "Execution time limits, cold starts"],
    ["Media files", "Blob Storage (Hot/Cool)", "Azure Files; VM disks", "Massive scale, SAS valet key, lifecycle tiers [5]", "Not queryable, so needs a metadata DB"],
    ["Media metadata", "Cosmos DB for NoSQL", "Table Storage; MongoDB on a VM", "Flexible JSON, horizontal partitions, autoscale RU/s [9]", "RU cost model; partition key fixed"],
    ["Users & relations", "Azure SQL (serverless)", "MySQL Flexible Server", "Integrity, joins, auto-pause when idle [10]", "Scales mainly vertically"],
  ];
  const tbl = [[hdr("Need"), hdr("Chosen"), hdr("Alternative"), hdr("Why chosen (pro)"), hdr("Trade-off (con)")]];
  rowsT.forEach((r, i) => {
    const f = { fill: { color: i % 2 ? C.background1 : C.background2 } };
    tbl.push([cell(r[0], Object.assign({ bold: true, color: C.text2 }, f)), cell(r[1], Object.assign({ bold: true, color: C.accent2 }, f)), cell(r[2], f), cell(r[3], f), cell(r[4], Object.assign({ color: C.accent1 }, f))]);
  });
  s.addTable(tbl, { x: 0.6, y: 1.35, w: 12.1, colW: [1.85, 2.45, 2.2, 3.35, 2.25], rowH: 0.62, border: { type: "solid", pt: 0.5, color: C.accent6 }, valign: "middle", margin: [4, 8, 4, 8] });
  s.addShape(S.ROUNDED_RECTANGLE, { x: 0.6, y: 6.05, w: 12.1, h: 0.75, rectRadius: 0.08, fill: { color: C.accent3, transparency: 88 }, line: { type: "none" } });
  T(s, [{ text: "Overall choice: ", options: { bold: true, color: C.text2 } }, { text: "serverless PaaS first (no VMs to patch or scale), keeping the brief's Logic Apps API and using Functions only where custom code is needed.", options: { color: C.text1 } }], { x: 0.85, y: 6.05, w: 11.7, h: 0.75, fontSize: 14, valign: "middle" });

  // =====================================================================
  // 11. ADVANCED FEATURES
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Critical evaluation" });
  s.addText("Advanced platform features planned for CW2", { placeholder: "title" });
  const adv = [
    ["FaChartLine", "accent3", "Application Insights + Azure Monitor", ["Distributed tracing across Logic Apps and Functions [14]", "Alerts on failed uploads, p95 latency and Cosmos DB RU throttling (429s)", "Dashboard for usage and cost trends"]],
    ["FaGithub", "text2", "CI/CD with GitHub Actions", ["On push to main: build → test → deploy static site, Functions and Logic App templates [15]", "Infrastructure as code (Bicep) for repeatable environments", "Manual approval gate for production"]],
    ["FaUserLock", "accent2", "Microsoft Entra External ID", ["Sign-up / sign-in with email or social accounts [12]", "OAuth 2.0 tokens checked on every API call", "Owner-only edit/delete; managed identities instead of keys"]],
    ["FaRobot", "accent1", "Azure AI Vision", ["Auto-tags and captions for images to power search [13]", "Flags adult or violent content, so media starts as 'pending'", "Video: analysis of the thumbnail frame"]],
  ];
  for (let i = 0; i < adv.length; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.15, y = 1.35 + row * 2.8;
    card(s, x, y, 5.95, 2.6);
    await iconCircle(s, adv[i][0], x + 0.25, y + 0.25, 0.7, adv[i][1]);
    T(s, adv[i][2], { x: x + 1.15, y: y + 0.25, w: 4.6, h: 0.7, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
    T(s, adv[i][3].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < adv[i][3].length - 1 } })), { x: x + 0.3, y: y + 1.05, w: 5.45, h: 1.45, fontSize: 14, paraSpaceAfter: 5 });
  }

  // =====================================================================
  // 12. SCALABILITY
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Critical evaluation" });
  s.addText("Assessment of scalability", { placeholder: "title" });
  const scl = [
    ["Edge", "Azure Front Door", "Global points of presence cache static files and thumbnails, offloading the origin [11]", "HORIZONTAL", "accent2"],
    ["API", "Logic Apps (Consumption)", "Multi-tenant: each request runs its own workflow instance, with no servers to size [7]", "AUTO-SCALE", "accent3"],
    ["Processing", "Functions (Flex)", "Instances added per event backlog and scale to zero; a max-instance cap controls cost [8]", "HORIZONTAL", "accent2"],
    ["Files", "Blob Storage", "Clients upload directly (Valet Key), removing the API bandwidth bottleneck [3], [5]", "PLATFORM", "accent3"],
    ["Metadata", "Cosmos DB", "Autoscale between 10% and 100% of max RU/s; /userId partitions spread load; add regions [9]", "HORIZONTAL", "accent2"],
    ["Relational", "Azure SQL serverless", "vCores scale between min/max and auto-pause; read replica for heavy reads [10]", "VERTICAL", "accent4"],
  ];
  T(s, [{ text: "TIER", options: {} }], { x: 0.6, y: 1.3, w: 1.3, h: 0.3, fontSize: 10, bold: true, color: C.accent5 });
  T(s, "RESOURCE", { x: 1.95, y: 1.3, w: 2.4, h: 0.3, fontSize: 10, bold: true, color: C.accent5 });
  T(s, "HOW IT SCALES", { x: 4.45, y: 1.3, w: 6.0, h: 0.3, fontSize: 10, bold: true, color: C.accent5 });
  T(s, "TYPE", { x: 10.85, y: 1.3, w: 1.85, h: 0.3, fontSize: 10, bold: true, color: C.accent5, align: "center" });
  scl.forEach((r, i) => {
    const y = 1.65 + i * 0.68;
    s.addShape(S.ROUNDED_RECTANGLE, { x: 0.6, y, w: 12.1, h: 0.6, rectRadius: 0.06, fill: { color: i % 2 ? C.background1 : C.background2 }, line: { type: "none" } });
    T(s, r[0], { x: 0.75, y, w: 1.2, h: 0.6, fontSize: 12, bold: true, color: C.accent5, valign: "middle" });
    T(s, r[1], { x: 1.95, y, w: 2.4, h: 0.6, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
    T(s, r[2], { x: 4.45, y, w: 6.25, h: 0.6, fontSize: 12, color: C.text1, valign: "middle" });
    s.addShape(S.ROUNDED_RECTANGLE, { x: 10.95, y: y + 0.13, w: 1.65, h: 0.34, rectRadius: 0.17, fill: { color: C[r[4]] }, line: { type: "none" } });
    T(s, r[3], { x: 10.95, y: y + 0.13, w: 1.65, h: 0.34, fontSize: 10, bold: true, color: C.background1, align: "center", valign: "middle" });
  });
  const keys = [["FaRoute", "Stateless API", "Any instance can serve any request; no sticky sessions."], ["FaBolt", "Async processing", "Heavy work runs after upload, off the request path."], ["FaLayerGroup", "Independent tiers", "Each tier scales on its own signal, so no single bottleneck."]];
  for (let i = 0; i < keys.length; i++) {
    const x = 0.6 + i * 4.1;
    await iconCircle(s, keys[i][0], x, 5.85, 0.55, "text2");
    T(s, [{ text: keys[i][1], options: { bold: true, fontSize: 13, color: C.text2, breakLine: true } }, { text: keys[i][2], options: { fontSize: 12, color: C.text1 } }], { x: x + 0.7, y: 5.78, w: 3.25, h: 0.95, valign: "middle" });
  }

  // =====================================================================
  // 13. LIMITATIONS
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Critical evaluation" });
  s.addText("Limitations and how to address them", { placeholder: "title" });
  T(s, "LIMITATION", { x: 0.6, y: 1.25, w: 5.6, h: 0.3, fontSize: 11, bold: true, color: C.accent1 });
  T(s, "MITIGATION", { x: 7.1, y: 1.25, w: 5.6, h: 0.3, fontSize: 11, bold: true, color: C.accent2 });
  const lim = [
    ["Logic Apps latency and timeouts", "Each connector adds latency, and synchronous responses are time-limited.", "Keep workflows short; cache GET feeds at Front Door; move hot paths to Functions."],
    ["Function cold starts", "The first thumbnail after an idle period is slow.", "Flex Consumption 'always ready' instances for the processing function [8]."],
    ["No video transcoding", "Azure Media Services retired in 2024 [16], so videos play at upload quality only.", "Enforce size/format limits; add an FFmpeg job (e.g. Container Apps) later."],
    ["Cosmos DB hot partitions", "A very active user can approach the 20 GB logical partition limit [9].", "Hierarchical partition key (/userId, /yearMonth)."],
    ["Public data endpoints", "Logic Apps Consumption cannot join a VNet, so data services rely on firewalls.", "Logic Apps Standard + private endpoints for production."],
    ["Single region, variable cost", "A region outage takes the service down; viral spikes raise the bill.", "Geo-redundant storage, Cosmos multi-region; budgets and autoscale caps."],
  ];
  lim.forEach((r, i) => {
    const y = 1.6 + i * 0.88;
    s.addShape(S.ROUNDED_RECTANGLE, { x: 0.6, y, w: 6.0, h: 0.78, rectRadius: 0.06, fill: { color: C.accent1, transparency: 91 }, line: { type: "none" } });
    T(s, [{ text: r[0], options: { bold: true, fontSize: 13, color: C.text2, breakLine: true } }, { text: r[1], options: { fontSize: 11, color: C.text1 } }], { x: 0.8, y, w: 5.7, h: 0.78, valign: "middle" });
    arrow(s, 6.68, y + 0.39, 7.02, y + 0.39, { width: 2, color: C.accent5 });
    s.addShape(S.ROUNDED_RECTANGLE, { x: 7.1, y, w: 5.6, h: 0.78, rectRadius: 0.06, fill: { color: C.accent2, transparency: 90 }, line: { type: "none" } });
    T(s, r[2], { x: 7.3, y, w: 5.3, h: 0.78, fontSize: 12, color: C.text1, valign: "middle" });
  });

  // =====================================================================
  // 14. CONCLUSION
  // =====================================================================
  pres.addSection({ title: "Conclusion" });
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Conclusion" });
  s.addText("Conclusion and future direction", { placeholder: "title" });
  T(s, "MediaVault combines static hosting, Logic Apps, Functions, Blob Storage, Cosmos DB and Azure SQL into a stateless, event-driven design where every tier scales on its own and costs follow real usage.", { x: 0.6, y: 1.25, w: 12.1, h: 0.85, fontSize: 16, color: C.accent6 });
  const concl = [
    ["FaLightbulb", "accent4", "Improvements", ["Private endpoints + Logic Apps Standard", "Video transcoding and adaptive streaming", "Multi-region active/active failover", "Redis cache for trending feeds"]],
    ["FaBriefcase", "accent2", "Commercial use", ["Freemium: storage quota tiers", "Blob lifecycle (Hot → Cool → Archive) protects margins", "Pay-as-you-go costs grow with revenue", "White-label for schools, events, photographers"]],
    ["FaBalanceScale", "accent1", "Ethical considerations", ["UK GDPR: data minimisation, UK South residency", "Right to erasure: DELETE removes blob + metadata", "AI tags can be biased, so humans review appeals", "Consent and copyright for people shown in media"]],
  ];
  for (let i = 0; i < concl.length; i++) {
    const x = 0.6 + i * 4.1;
    s.addShape(S.ROUNDED_RECTANGLE, { x, y: 2.35, w: 3.9, h: 4.1, rectRadius: 0.1, fill: { color: C.background1, transparency: 92 }, line: { color: C.accent5, width: 0.75 } });
    await iconCircle(s, concl[i][0], x + 0.25, 2.6, 0.65, concl[i][1]);
    T(s, concl[i][2], { x: x + 1.05, y: 2.6, w: 2.7, h: 0.65, fontSize: 18, bold: true, color: C.background1, valign: "middle" });
    T(s, concl[i][3].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < concl[i][3].length - 1 } })), { x: x + 0.25, y: 3.5, w: 3.45, h: 3.1, fontSize: 14, color: C.background2, paraSpaceAfter: 8 });
  }

  // =====================================================================
  // 15. REFERENCES
  // =====================================================================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Conclusion" });
  s.addText("References", { placeholder: "title" });
  const refs = [
    "P. Mell and T. Grance, \"The NIST definition of cloud computing,\" NIST SP 800-145, 2011.",
    "Microsoft, \"Cloud design patterns,\" Azure Architecture Center. learn.microsoft.com/azure/architecture/patterns",
    "Microsoft, \"Valet Key pattern.\" learn.microsoft.com/azure/architecture/patterns/valet-key",
    "Microsoft, \"Azure Well-Architected Framework.\" learn.microsoft.com/azure/well-architected",
    "Microsoft, \"Introduction to Azure Blob Storage.\" learn.microsoft.com/azure/storage/blobs/storage-blobs-introduction",
    "Microsoft, \"Static website hosting in Azure Storage.\" learn.microsoft.com/azure/storage/blobs/storage-blob-static-website",
    "Microsoft, \"What is Azure Logic Apps?\" learn.microsoft.com/azure/logic-apps/logic-apps-overview",
    "Microsoft, \"Azure Functions Flex Consumption plan hosting.\" learn.microsoft.com/azure/azure-functions/flex-consumption-plan",
    "Microsoft, \"Partitioning and horizontal scaling in Azure Cosmos DB.\" learn.microsoft.com/azure/cosmos-db/partitioning-overview",
    "Microsoft, \"Azure SQL Database serverless.\" learn.microsoft.com/azure/azure-sql/database/serverless-tier-overview",
    "Microsoft, \"What is Azure Front Door?\" learn.microsoft.com/azure/frontdoor/front-door-overview",
    "Microsoft, \"Overview of Microsoft Entra External ID.\" learn.microsoft.com/entra/external-id/external-identities-overview",
    "Microsoft, \"What is Azure AI Vision?\" learn.microsoft.com/azure/ai-services/computer-vision/overview",
    "Microsoft, \"Application Insights overview.\" learn.microsoft.com/azure/azure-monitor/app/app-insights-overview",
    "GitHub, \"Deploying to Azure.\" docs.github.com/actions/deployment/deploying-to-your-cloud-provider/deploying-to-azure",
    "Microsoft, \"Azure Media Services retirement guide.\" learn.microsoft.com/azure/media-services/latest/azure-media-services-retirement",
  ];
  const refRuns = (arr, start) => arr.flatMap((r, i) => [
    { text: `[${start + i}] `, options: { bold: true, color: C.text2 } },
    { text: r, options: { color: C.text1, breakLine: i < arr.length - 1 } },
  ]);
  T(s, refRuns(refs.slice(0, 8), 1), { x: 0.6, y: 1.3, w: 5.95, h: 4.55, fontSize: 12, paraSpaceAfter: 6 });
  T(s, refRuns(refs.slice(8), 9), { x: 6.75, y: 1.3, w: 5.95, h: 4.55, fontSize: 12, paraSpaceAfter: 6 });
  T(s, "All web sources accessed November 2026. Module lecture and lab materials (COM682, Weeks 3–5) also informed the design.", { x: 0.6, y: 5.9, w: 12.1, h: 0.3, fontSize: 10, italic: true, color: C.accent5 });
  s.addShape(S.ROUNDED_RECTANGLE, { x: 0.6, y: 6.25, w: 12.1, h: 0.62, rectRadius: 0.06, fill: { color: C.accent4, transparency: 85 }, line: { type: "none" } });
  T(s, [{ text: "AI use acknowledgement: ", options: { bold: true, color: C.text2 } }, { text: "Claude (Anthropic, 2026) was used to help draft slide text and diagrams. All design decisions were reviewed, edited and verified by the author.", options: { color: C.text1 } }], { x: 0.8, y: 6.25, w: 11.7, h: 0.62, fontSize: 12, valign: "middle" });

  // =====================================================================
  // 16. CHECKLIST (not counted)
  // =====================================================================
  pres.addSection({ title: "Appendix" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Appendix" });
  s.addText("Appendix III: Submission checklist", { placeholder: "title" });
  T(s, "Not counted in the 15-slide limit", { x: 0.6, y: 1.1, w: 6, h: 0.3, fontSize: 12, italic: true, color: C.accent5 });
  const chk = [
    ["General", [["Title slide: project name, description, name, student number", "1"], ["PowerPoint format, within 15 slides (excl. references)", "1–15"]]],
    ["Problem & context", [["Problem discussed", "2"], ["Scalability issues of resources identified", "3"]]],
    ["Solution design", [["Solution architecture diagram (Azure resources only)", "4"], ["Multimedia hosting webpage layout (wireframes)", "6"], ["Network/architecture diagram: compute, storage, networking", "5"], ["Database schema: resource list + ERD (+ NoSQL JSON)", "7–8"], ["REST API for CRUD in a UML diagram", "9"], ["Static HTML hosting · REST endpoints · SQL + NoSQL", "4, 7–9"]]],
    ["Critical evaluation", [["Overview of advanced features", "11"], ["Assessment of limitations", "13"], ["Assessment of scalability (horizontal/vertical, autoscale)", "12"], ["Justification of Azure services vs alternatives", "10"]]],
    ["Conclusion & references", [["Concluding comments", "14"], ["References", "15"]]],
  ];
  const colsChk = [[chk[0], chk[1], chk[2]], [chk[3], chk[4]]];
  for (let c = 0; c < 2; c++) {
    let y = 1.5; const x = 0.6 + c * 6.2;
    for (const [grp, items] of colsChk[c]) {
      T(s, grp, { x, y, w: 5.9, h: 0.3, fontSize: 13, bold: true, color: C.text2 }); y += 0.33;
      for (const [txt, sl] of items) {
        s.addShape(S.ROUNDED_RECTANGLE, { x, y: y + 0.03, w: 0.24, h: 0.24, rectRadius: 0.04, fill: { color: C.accent2 }, line: { type: "none" } });
        T(s, "✓", { x, y: y + 0.03, w: 0.24, h: 0.24, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
        T(s, txt, { x: x + 0.35, y, w: 4.7, h: 0.3, fontSize: 11, color: C.text1, valign: "middle" });
        T(s, "Slide " + sl, { x: x + 5.05, y, w: 0.9, h: 0.3, fontSize: 10, color: C.accent5, align: "right", valign: "middle" });
        y += 0.32;
      }
      y += 0.12;
    }
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
