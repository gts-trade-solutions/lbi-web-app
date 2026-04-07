/**
 * lib/download.ts (FULL FILE)
 *
 * ✅ Fixes included in this version:
 * ✅ Single STYLE token system for fonts/spacing/margins (consistent look everywhere)
 * ✅ Table width matches column widths (prevents stretching/clipping)
 * ✅ Columns rebalanced: DETAILS reduced, PHOTO widened (full image visible)
 * ✅ Row height controlled via STYLE tokens so images never get cut
 * ✅ Photo sizing adjusted for the PHOTO cell width (single vs multi)
 * ✅ Alignment kept clean (centered headers, consistent cell padding)
 */

import JSZip from "jszip";
import sharp from "sharp";
import fs from "fs/promises";
import path from "path";
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  HeightRule,
  ImageRun,
  Packer,
  PageNumber,
  PageOrientation,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  UnderlineType,
  VerticalAlign,
  WidthType,
  TextWrappingType,
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  VerticalPositionAlign,
  VerticalPositionRelativeFrom,
} from "docx";


function bytesFromBase64(b64: string): Uint8Array {
  return new Uint8Array(Buffer.from(b64, "base64"));
}

async function readLocalPublicFileBytes(publicPath: string): Promise<Uint8Array | null> {
  try {
    const clean = String(publicPath || "").split("?")[0].split("#")[0];
    const rel = clean.replace(/^\/+/, "");
    const full = path.join(process.cwd(), "public", rel);
    const buf = await fs.readFile(full);
    return new Uint8Array(buf);
  } catch {
    return null;
  }
}

function maybeAbsoluteUrl(u: string): string {
  const raw = String(u || "").trim();
  if (!raw) return raw;
  if (/^https?:\/\//i.test(raw) || raw.startsWith("data:")) return raw;
  const base = process.env.NEXT_PUBLIC_SITE_URL || process.env.NEXT_PUBLIC_APP_URL || process.env.SITE_URL || "";
  if (base && raw.startsWith("/")) return `${base.replace(/\/+$/, "")}${raw}`;
  return raw;
}

/** =========================
 * ✅ STYLE TOKENS (single source of truth)
 * ========================= */
const STYLE = {
  font: {
    cell: 28, // ~11pt
    cellSmall: 28, // ~9pt
    header: 28, // ~12pt
    // NOTE: docx TextRun.size is in half-points (pt * 2).
    // User-required: all section/page titles = 24pt => 48 half-points.
    sectionTitle: 48,
    // Cover page title sizes are handled explicitly (46pt / 36pt / 24pt).
    title: 56, // legacy/unused for cover main title
    meta: 32, // footer/meta
  },
  spacing: {
    none: { before: 0, after: 0 },
    cell: { before: 40, after: 40, line: 320 },
    cellTight: { before: 0, after: 0, line: 276 },
    section: { before: 120, after: 120, line: 360 },
    sectionTitle: { before: 120, after: 60 },
  },
  cellMargins: { top: 80, bottom: 80, left: 120, right: 120 } as any,
  row: { height: 3800 },
  photo: {
    single: { w: 210, h: 120 },
    multi: { w: 145, h: 90 },
  },
};


/** =========================
 * ✅ Reliable page border patch (matches client reference)
 * Adds <w:pgBorders ...> to every section in word/document.xml
 * ========================= */
// NOTE: Use offsetFrom="page" so borders are always visible on all sides.
// Content is kept inside the border by using slightly larger page margins.
const PG_BORDERS_XML =
  '<w:pgBorders w:offsetFrom="page">' +
  '<w:top w:val="double" w:sz="6" w:space="48" w:color="FF0000"/>' +
  '<w:left w:val="double" w:sz="6" w:space="48" w:color="FF0000"/>' +
  '<w:bottom w:val="double" w:sz="6" w:space="48" w:color="FF0000"/>' +
  '<w:right w:val="double" w:sz="6" w:space="48" w:color="FF0000"/>' +
  "</w:pgBorders>";

// ✅ Rounded input box background (PNG) to mimic border-radius in DOCX (Word table cells have no real borderRadius)
const __ROUNDED_INPUT_PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAA4QAAAB4CAYAAACq9jzEAAAEUUlEQVR4nO3dPVLrMBiG0XD3BBUUsHAooIJFQcXMnZDYkiXZjt9z2kxst898+jmdAAAAAAAAAAAAAAAAAIAjulvzZa/vn99rvg8AAOAWvTw9rNJqQ18iAAEAANqNCsQhDxWCAAAA/fUOw64PE4IAAADj9QrDLg8RggAAAOtrDcN/rR8gBgEAALbR2mNNNbnk5c+P9y2vBAAAOLS3j6/q/yydFC76U00ICkAAAIDlagKxNgyrg7A0BoUgAABAP6VhWBOFVUFYEoNCEAAAYJySMCyNwuZDZf4nBgEAAMbq2V3FQTg3HRSDAAAA65jrr9KtfkVjxKmHCUEAAIDtTC0hnVs6OjshdM8gAADAbZrruaY9hKaDAAAA22rpsskgtFQUAABg/6b6bKrrFk0IxSAAAMC+LOm0q0Fo7yAAAMAxXOu76gmh6SAAAMA+1fbaxSA0HQQAADiWS51XNSE0HQQAANi3mm5runYCAACA2yUIAQAAQv0JQvsHAQAAjum894onhPYPAgAA3IbSfrNkFAAAIJQgBAAACCUIAQAAQglCAACAUIIQAAAglCAEAAAIJQgBAABCCUIAAIBQghAAACCUIAQAAAglCAEAAEIJQgAAgFCCEAAAIJQgBAAACCUIAQAAQglCAACAUIIQAAAglCAEAAAIJQgBAABCCUIAAIBQghAAACCUIAQAAAglCAEAAEIJQgAAgFCCEAAAIJQgBAAACCUIAQAAQglCAACAUIIQAAAglCAEAAAIJQgBAABCCUIAAIBQghAAACCUIAQAAAglCAEAAEIJQgAAgFCCEAAAIJQgBAAACCUIAQAAQglCAACAUIIQAAAglCAEAAAIJQgBAABCCUIAAIBQghAAACCUIAQAAAglCAEAAEIJQgAAgFCCEAAAIJQgBAAACCUIAQAAQglCAACAUIIQAAAglCAEAAAIJQgBAABCCUIAAIBQghAAACCUIAQAAAglCAEAAEIJQgAAgFCCEAAAIJQgBAAACCUIAQAAQglCAACAUMVB+PbxNfI7AAAA6KS03/4E4cvTw133rwEAAGBz571nySgAAEAoQQgAABCqKgjtIwQAANi3mm67GIT2EQIAABzLpc6rXjJqSggAALBPtb12NQhNCQEAAI7hWt8tOlTGlBAAAGBflnTaZBBOTQlFIQAAwD5M9dlU1zVdOyEKAQAAttXSZbNBaC8hAADAbZrrueLYe33//J76/fnxvvRRAAAANJqbDJYM94qXjM49zPJRAACAdfSIwdOpcQ/hOVEIAAAwVs/uqt4fOLd09JclpAAAAP2UhmDNOTCLDowpjcLTSRgCAAC0qJkI1h4K2nSCaE0Y/hKIAAAA1y1ZErr0dojmKyWWRCEAAAB9tFwV2HyojHsKAQAAttHaY11jzrQQAABgvF6DuSHTPWEIAADQX+8VmkOXewpDAACAdqO26q26/08gAgAAzHNWCwAAAAAAAAAAAACd/ADtkM/Q/w16FgAAAABJRU5ErkJggg==";
function __roundedInputPngBytes() {
  return bytesFromBase64(__ROUNDED_INPUT_PNG_BASE64);
}


async function applyRedPageBordersToDocxBytes(input: ArrayBuffer | Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(input as any);
  const docXmlFile = zip.file("word/document.xml");
  if (!docXmlFile) {
    // if structure unexpected, return original bytes
    const out = input instanceof Uint8Array ? input : new Uint8Array(input);
    return out;
  }

  let xml = await docXmlFile.async("string");

  // Remove any existing pgBorders then insert ours into every sectPr
  xml = xml.replace(/<w:pgBorders[\s\S]*?<\/w:pgBorders>/g, "");
  xml = xml.replace(/<\/w:sectPr>/g, PG_BORDERS_XML + "</w:sectPr>");

  zip.file("word/document.xml", xml);

  const outBytes = await zip.generateAsync({ type: "uint8array" });
  return outBytes;
}


/** =========================
 * GPX Types
 * ========================= */
type GPXPoint = { lat: number; lon: number; time?: string };
function isoUtc(d: Date) {
  return d.toISOString();
}

/** =========================
 * Parses strings like:
 * "N28 02.912 E84 48.869"
 * "N28.12345 E84.98765"
 * "N28 02.912\nE84 48.869"
 * ========================= */
function parseNEToDecimal(ne: string): { lat: number; lon: number } | null {
  const t = String(ne || "")
    .replace(/\r/g, " ")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!t) return null;

  // DM format: N12 53.397 E79 54.775
  let m = t.match(
    /^([NS])\s*(\d{1,3})\s+(\d{1,2}(?:\.\d+)?)\s+([EW])\s*(\d{1,3})\s+(\d{1,2}(?:\.\d+)?)$/i
  );
  if (m) {
    const ns = m[1].toUpperCase();
    const latDeg = Number(m[2]);
    const latMin = Number(m[3]);
    const ew = m[4].toUpperCase();
    const lonDeg = Number(m[5]);
    const lonMin = Number(m[6]);

    let lat = latDeg + latMin / 60;
    let lon = lonDeg + lonMin / 60;

    if (ns === "S") lat = -lat;
    if (ew === "W") lon = -lon;

    return { lat, lon };
  }

  // Decimal format: N12.913786 E79.856013
  m = t.match(
    /^([NS])\s*(\d{1,3}(?:\.\d+)?)\s+([EW])\s*(\d{1,3}(?:\.\d+)?)$/i
  );
  if (m) {
    const ns = m[1].toUpperCase();
    let lat = Number(m[2]);
    const ew = m[3].toUpperCase();
    let lon = Number(m[4]);

    if (ns === "S") lat = -lat;
    if (ew === "W") lon = -lon;

    return { lat, lon };
  }

  return null;
}

/** =========================
 * ✅ Reverse Geocode (OSM Nominatim)
 * ========================= */
const REVERSE_CACHE = new Map<string, string>();
const REVERSE_INFLIGHT = new Map<string, Promise<string>>();
const REVERSE_TIMEOUT_MS = 9000;

function coordKey(lat: number, lon: number) {
  const r = (n: number) => Math.round(n * 100000) / 100000; // ~1m for better deep-location accuracy
  return `${r(lat)},${r(lon)}`;
}

function pickFirst(obj: any, keys: string[]) {
  for (const k of keys) {
    const v = obj?.[k];
    if (typeof v === "string" && v.trim()) return v.trim();
  }
  return "";
}

function compactUniqueParts(parts: string[]) {
  const out: string[] = [];
  const seen = new Set<string>();
  for (const p of parts) {
    const t = String(p || "").trim();
    if (!t) continue;
    const key = t.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(t);
  }
  return out;
}

function formatOsmAddress(addr: any) {
  const house = pickFirst(addr, ["house_number"]);
  const road = pickFirst(addr, [
    "road",
    "pedestrian",
    "residential",
    "service",
    "footway",
    "path",
    "cycleway",
  ]);
  const landmark = pickFirst(addr, [
    "building",
    "amenity",
    "shop",
    "office",
    "industrial",
    "commercial",
    "tourism",
    "bridge",
    "tunnel",
    "man_made",
    "railway",
  ]);
  const p1 = pickFirst(addr, ["neighbourhood", "suburb", "quarter", "hamlet", "locality"]);
  const p2 = pickFirst(addr, ["city_district", "district", "borough", "county", "state_district"]);
  const p3 = pickFirst(addr, ["city", "town", "village", "municipality"]);
  const p4 = pickFirst(addr, ["state"]);

  const streetLine = [house, road].filter(Boolean).join(" ").trim();
  const parts = compactUniqueParts([streetLine, landmark, p1, p2, p3, p4]);
  return parts.slice(0, 6).join(", ");
}

async function fetchWithTimeout(url: string, ms: number) {
  // Disabled for client-side DOCX export to avoid browser CORS/network failures.
  // Kept only for compatibility; returns a rejected promise if called accidentally.
  throw new Error("Reverse geocoding disabled in client export");
}

async function reverseGeocodeOSM(lat: number, lon: number): Promise<string> {
  // Browser export should not call external reverse-geocoding APIs.
  // Use stored location when available; otherwise fall back to coordinates.
  return "";
}

/** =========================
 * GPX Generator
 * ========================= */
function haversineKm(lat1: number, lon1: number, lat2: number, lon2: number) {
  const R = 6371;
  const toRad = (x: number) => (x * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) *
      Math.cos(toRad(lat2)) *
      Math.sin(dLon / 2) *
      Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function formatDecimalToDM(value: number, kind: "lat" | "lon") {
  const abs = Math.abs(value);
  const deg = Math.floor(abs);
  const min = (abs - deg) * 60;

  const dir =
    kind === "lat"
      ? value >= 0
        ? "N"
        : "S"
      : value >= 0
        ? "E"
        : "W";

  return `${dir}${deg} ${min.toFixed(3)}`;
}

function formatNEFromLatLon(lat: number, lon: number) {
  return `${formatDecimalToDM(lat, "lat")}
${formatDecimalToDM(lon, "lon")}`;
}

function fallbackLocationFromLatLon(lat: number, lon: number) {
  return `${lat.toFixed(6)}, ${lon.toFixed(6)}`;
}

function pointSortValue(raw: any, index: number) {
  const candidates = [
    raw.point_key,
    raw.gps_no,
    raw.gps,
    raw.no,
    raw.seq,
    raw.sequence,
    raw.point_no,
    raw.point_index,
    raw.idx,
    raw.index,
  ];

  for (const c of candidates) {
    const n = Number(c);
    if (Number.isFinite(n)) return n;
  }

  return index + 1;
}

function toGpxXml(params: { name: string; creator?: string; points: GPXPoint[] }) {
  const creator = params.creator || "Recorded in TSPL Web App";
  const name = (params.name || "Export").trim() || "Export";

  const pts = params.points || [];
  const now = new Date();

  const startTime = pts[0]?.time || isoUtc(now);
  const endTime = pts[pts.length - 1]?.time || startTime;

  let lengthKm = 0;
  for (let i = 1; i < pts.length; i++) {
    lengthKm += haversineKm(pts[i - 1].lat, pts[i - 1].lon, pts[i].lat, pts[i].lon);
  }

  const durationMs = (() => {
    try {
      const a = Date.parse(startTime);
      const b = Date.parse(endTime);
      if (Number.isFinite(a) && Number.isFinite(b) && b >= a) return b - a;
    } catch {}
    return 0;
  })();

  const trkptsXml = pts
    .map((p) => {
      const timeXml = p.time ? `\n        <time>${p.time}</time>` : "";
      return `      <trkpt lat="${p.lat}" lon="${p.lon}">${timeXml}\n      </trkpt>`;
    })
    .join("\n");

  return `<?xml version='1.0' encoding='UTF-8' standalone='yes' ?>
<gpx xmlns="http://www.topografix.com/GPX/1/1"
     xmlns:geotracker="http://ilyabogdanovich.com/gpx/extensions/geotracker"
     xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
     xsi:schemaLocation="http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd"
     version="1.1"
     creator="${creator}">
  <metadata>
    <name>${name}</name>
    <time>${isoUtc(new Date())}</time>
  </metadata>
  <trk>
    <name>${name}</name>
    <src>${creator}</src>
    <extensions>
      <geotracker:meta>
        <geotracker:length>${lengthKm.toFixed(2)}</geotracker:length>
        <geotracker:duration>${durationMs}</geotracker:duration>
        <geotracker:creationtime>${startTime}</geotracker:creationtime>
        <geotracker:activity>0</geotracker:activity>
      </geotracker:meta>
    </extensions>
    <trkseg>
${trkptsXml}
    </trkseg>
  </trk>
</gpx>`;
}

/** =========================
 * TSPL FORMAT SETTINGS
 * ========================= */

// ✅ IMPORTANT: table width should match the column widths sum (prevents Word stretching & clipping)
const TABLE_TOTAL_W = 23578;

// ✅ Rebalanced widths (sum = 15848) — PHOTO gets more width; DETAILS reduced
// Columns: [GPS, KMS, NE, D1, D2, L1, L2, DESC, MOVE, P1, P2, P3]
const GRID_COLS = [1637, 1934, 2827, 1562, 1562, 1934, 1934, 1785, 1338, 2379, 2379, 2307];

const HEADER_FILL = "365F91";
const PHOTO_PAGE_HEADER_FILL = "4CAF50";
const PHOTO_PAGE_ROW_FILL = "D9EAD3";

const PHOTO_THEME = {
  green: {
    header: "4CAF50",
    body: "D9EAD3",
    text: "0B3D2E",
  },
  yellow: {
    header: "FFFF00",
    body: "FEFECA",
    text: "7A5D00",
  },
  red: {
    header: "FF0000",
    body: "E57373",
    text: "7A1C1C",
  },
  default: {
    header: "4CAF50",
    body: "D9EAD3",
    text: "0B3D2E",
  },
};

const PAGE_BORDER_COLOR = "C00000";

const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const CELL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
// ✅ Photo cells: remove the two horizontal lines (top/bottom) so images look clean
const PHOTO_CELL_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: BORDER, right: BORDER };

const DEBUG_PHOTOS = false;

// ✅ A4 size (TWIPS)
const A4_W = 11906;
const A4_H = 16838;

// ✅ A3 size (TWIPS) (11.69" x 16.54")
// NOTE: Word uses twips (1 inch = 1440 twips).
// A3 = 11.69 x 16.54 in  -> 16838 x 23811 twips (commonly used rounded values)
// We intentionally use A3 for the wide, reference-like landscape pages.
const A3_W = 16838;
const A3_H = 23811;

// ✅ Tight margins
const COVER_MARGIN = {
  top: 240,
  bottom: 240,
  left: 520, right: 520,
  header: 240,
  footer: 320,
  gutter: 0,
};

const TABLE_MARGIN = {
  top: 240,
  bottom: 240,
  left: 520, right: 520,
  header: 420,
  footer: 520,
  gutter: 0,
};

/** ✅ Watermark opts */
export type WatermarkOptions = { enabled?: boolean; text?: string };

export type CoverOptions = {
  enabled?: boolean;
  logoUrl?: string;
  logoWidth?: number;
  logoHeight?: number;
  rightTopText?: string;
  topCenterText?: string;
  recommendationText?: string;
  footerLeftText?: string;
  footerEmail?: string;
  footerWebsite?: string;
  datedLabel?: string;
  date?: string | Date;
};

export type DownloadOpts = {
  includePhotos?: boolean;
  fileName?: string;
  watermark?: WatermarkOptions;
  cover?: CoverOptions;
};

/** =========================
 * ✅ ROUTE SURVEY (Objective + Route Map + GA Drawing)
 * Pulls from:
 *  - project_route_pages (latest)
 *  - project_route_page_images (latest/first)
 * ========================= */
type ProjectRoutePageRow = {
  id: string;
  project_id: string;
  objective?: string | null;
  map_mode?: string | null; // 'preset' | 'upload'
  preset_map_key?: string | null; // ex: 'route1.png' OR full url OR '/maps/route1.png'
  map_file_url?: string | null; // full public url if uploaded
  created_at?: string | null;
};

type ProjectRouteImageRow = {
  id: string;
  project_id: string;
  project_page_id: string;
  file_url: string;
  created_at?: string | null;
};

function isHttpUrl(u: string) {
  return /^https?:\/\//i.test(u);
}

function resolvePresetMapUrl(presetKey: string) {
  const t = String(presetKey || "").trim();
  if (!t) return "";
  if (isHttpUrl(t) || t.startsWith("data:")) return t;
  // If you saved only a file name like "route1.png" -> assume /public/maps/
  if (!t.includes("/")) return `/maps/${t}`;
  // If you saved "maps/route1.png"
  if (!t.startsWith("/")) return `/${t}`;
  return t;
}

async function getProjectRouteSetup(
  supabase: any,
  projectId: string
): Promise<{ objective: string; routeMapUrl: string; gaImageUrls: string[]; locations: string[] } | null> {
  const { data: page, error: pErr } = await supabase
    .from("project_route_pages")
    .select("id, project_id, objective, map_mode, preset_map_key, map_file_url, created_at")
    .eq("project_id", projectId)
    .order("created_at", { ascending: false })
    .limit(1)
    .maybeSingle();

  if (pErr) throw pErr;
  if (!page?.id) return null;

  const pr = page as ProjectRoutePageRow;
  const routeMapUrl =
    (pr.map_file_url && String(pr.map_file_url).trim()) ||
    (pr.preset_map_key ? resolvePresetMapUrl(pr.preset_map_key) : "") ||
    "";

  const { data: imgs, error: iErr } = await supabase
    .from("project_route_page_images")
    .select("id, project_id, project_page_id, file_url, created_at")
    .eq("project_id", projectId)
    .eq("project_page_id", pr.id)
    .order("created_at", { ascending: true })
    .limit(20);

  if (iErr) throw iErr;

  const gaImageUrls = (imgs || [])
    .map((x: any) => String((x as ProjectRouteImageRow)?.file_url || "").trim())
    .filter(Boolean);

  // ✅ locations (from project_route_page_locations)
  // Fetch 4 labels for THIS project (project_id), ordered by sort_order ASC.
  let locations: string[] = [];
  const { data: locs, error: lErr } = await supabase
    .from("project_route_page_locations")
    .select("label, sort_order")
    .eq("project_id", projectId)
    .eq("project_page_id", pr.id)
    .order("sort_order", { ascending: true })
    .limit(4);

  if (lErr) throw lErr;
  locations = (locs || []).map((x: any) => String(x?.label || "").trim()).filter(Boolean);

  return { objective: String(pr.objective || "").trim(), routeMapUrl, gaImageUrls, locations };
}

function bodyText(text: string) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: STYLE.spacing.section,
    children: [new TextRun({ text: (text || "—").trim() || "—", size: STYLE.font.cell })],
  });
}

function objectiveText(text: string) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: STYLE.spacing.section,
    // User-required: objective content = 24pt => 48 half-points
    children: [new TextRun({ text: (text || "—").trim() || "—", size: 48 })],
  });
}

function underlineLabel(text: string, size: number = STYLE.font.sectionTitle) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: STYLE.spacing.sectionTitle,
    children: [
      new TextRun({
        text,
        bold: true,
        size,
        underline: { type: UnderlineType.SINGLE },
      }),
    ],
  });
}

function centeredImage(bytes: Uint8Array, w: number, h: number) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 120 },
    children: [new ImageRun({ data: bytes, transformation: { width: w, height: h } })],
  });
}



function getImageDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  if (!bytes || bytes.length < 24) return null;

  // PNG: signature + IHDR (width/height at offsets 16/20)
  const isPng =
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a;

  if (isPng && bytes.length >= 24) {
    const w = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
    const h = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
    if (w > 0 && h > 0) return { width: w >>> 0, height: h >>> 0 };
  }

  // JPEG: scan for SOF marker that contains width/height
  const isJpg = bytes[0] === 0xff && bytes[1] === 0xd8;
  if (isJpg) {
    let i = 2;
    while (i + 9 < bytes.length) {
      if (bytes[i] !== 0xff) {
        i++;
        continue;
      }
      // Skip fill bytes
      while (i < bytes.length && bytes[i] === 0xff) i++;
      if (i >= bytes.length) break;
      const marker = bytes[i];
      i++;

      // Standalone markers (no length)
      if (marker === 0xd8 || marker === 0xd9) continue;
      if (marker >= 0xd0 && marker <= 0xd7) continue;

      if (i + 1 >= bytes.length) break;
      const segLen = (bytes[i] << 8) | bytes[i + 1];
      if (segLen < 2 || i + segLen - 2 >= bytes.length) break;

      const isSOF =
        (marker >= 0xc0 && marker <= 0xc3) ||
        (marker >= 0xc5 && marker <= 0xc7) ||
        (marker >= 0xc9 && marker <= 0xcb) ||
        (marker >= 0xcd && marker <= 0xcf);

      if (isSOF) {
        // SOF: [lenHi lenLo][precision][heightHi heightLo][widthHi widthLo]...
        const p = i + 2;
        const h = (bytes[p + 1] << 8) | bytes[p + 2];
        const w = (bytes[p + 3] << 8) | bytes[p + 4];
        if (w > 0 && h > 0) return { width: w, height: h };
        break;
      }

      i += segLen;
    }
  }

  return null;
}


function fitTransform(bytes: Uint8Array, maxW: number, maxH: number) {
  const dim = getImageDimensions(bytes);
  let w = maxW;
  let h = maxH;

  if (dim && dim.width > 0 && dim.height > 0) {
    const scale = Math.min(maxW / dim.width, maxH / dim.height);
    w = Math.max(1, Math.round(dim.width * scale));
    h = Math.max(1, Math.round(dim.height * scale));
  }
  return { width: w, height: h };
}

function centeredImageFit(bytes: Uint8Array, maxW: number, maxH: number) {
  const { width: w, height: h } = fitTransform(bytes, maxW, maxH);

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 120 },
    children: [new ImageRun({ data: bytes, transformation: { width: w, height: h } })],
  });
}


// ✅ Trim large white margins from GA drawings so the actual drawing is centered visually.
const __TRIM_CACHE = new Map<string, Uint8Array>();

async function trimWhiteMarginsToPng(bytes: Uint8Array): Promise<Uint8Array> {
  try {
    const key = `${bytes.length}:${bytes[0]}:${bytes[1]}:${bytes[2]}:${bytes[3]}`;
    const cached = __TRIM_CACHE.get(key);
    if (cached) return cached;

    const out = await sharp(Buffer.from(bytes)).trim().png().toBuffer();
    const outBytes = new Uint8Array(out);
    __TRIM_CACHE.set(key, outBytes);
    return outBytes;
  } catch {
    return bytes;
  }
}

async function centeredImageFitTrim(bytes: Uint8Array, maxW: number, maxH: number) {
  const trimmed = await trimWhiteMarginsToPng(bytes);
  const { width: w, height: h } = fitTransform(trimmed, maxW, maxH);

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 120 },
    children: [new ImageRun({ data: trimmed, transformation: { width: w, height: h } })],
  });
}

function sectionPropsA3Landscape() {
  return {
    page: {
      size: { width: A3_W, height: A3_H, orientation: PageOrientation.LANDSCAPE },
      margin: TABLE_MARGIN as any,
      ...(pageBordersTSPL() as any),
    } as any,
  };
}

function sectionPropsA4Portrait() {
  return {
    page: {
      size: { width: A4_W, height: A4_H, orientation: PageOrientation.PORTRAIT },
      margin: TABLE_MARGIN as any,
      ...(pageBordersTSPL() as any),
    } as any,
  };
}

async function bytesFromUrlForDocx(url: string): Promise<Uint8Array | null> {
  const u = String(url || "").trim();
  if (!u) return null;
  if (u.startsWith("/")) return await readLocalPublicFileBytes(u);
  try {
    return await fetchBytes(maybeAbsoluteUrl(u));
  } catch {
    return null;
  }
}

async function buildObjectiveRouteMapSection(params: {
  projectName: string;
  objective: string;
  routeMapBytes: Uint8Array | null;
  routeLocations?: string[];
  footerDate?: string | Date;
}) {
  const { projectName, objective, routeMapBytes, routeLocations, footerDate } = params;

  const children: Paragraph[] = [];
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [
        new TextRun({
          text: String(projectName || "").toUpperCase(),
          bold: true,
          size: STYLE.font.sectionTitle,
          color: "667085",
        }),
      ],
    })
  );

  // User-required: OBJECTIVE label = 24pt => 48 half-points
  children.push(underlineLabel("OBJECTIVE:", 48));
  children.push(objectiveText(objective || "—"));

  // User-required: ROUTE MAP label = 24pt => 48 half-points
  children.push(underlineLabel("ROUTE MAP:", 48));
  children.push(new Paragraph({ spacing: { before: 0, after: 0 }, text: "" }));

if (routeMapBytes) {
    // ✅ Show location boxes only when at least one location has value.
    // Otherwise show only the map centered.
    const rawLocs = (routeLocations || []).slice(0, 4);
    const hasAnyLocation = rawLocs.some((x) => String(x || "").trim());

    const mapImgPara = centeredImageFit(routeMapBytes, 1100, 700);

    if (!hasAnyLocation) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 0 },
          children: [],
        })
      );
      (children as any).push(mapImgPara);
    } else {
      const locs = [...rawLocs];
      while (locs.length < 4) locs.push("");

      const leftRows = locs.map((label, idx) => {
        const iconText = idx === 3 ? "📍" : "○";
        const iconColor = idx === 3 ? "B42318" : "101828";

        return new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: 20, type: WidthType.PERCENTAGE },
              margins: { top: 90, bottom: 90, left: 120, right: 120 } as any,
              verticalAlign: VerticalAlign.CENTER,
              borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
              },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun({ text: iconText, size: 36, color: iconColor, bold: true })],
                }),
              ],
            }),
            new TableCell({
              width: { size: 80, type: WidthType.PERCENTAGE },
              margins: { top: 90, bottom: 90, left: 220, right: 220 } as any,
              verticalAlign: VerticalAlign.CENTER,
              borders: {
                top: { style: BorderStyle.SINGLE, size: 6, color: "D0D5DD" },
                bottom: { style: BorderStyle.SINGLE, size: 6, color: "D0D5DD" },
                left: { style: BorderStyle.SINGLE, size: 6, color: "D0D5DD" },
                right: { style: BorderStyle.SINGLE, size: 6, color: "D0D5DD" },
              },
              children: [
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [
                    new TextRun({
                      text: (label || "").trim() || " ",
                      size: 32,
                      color: "101828",
                      bold: true,
                    }),
                  ],
                }),
              ],
            }),
          ],
        });
      });

      const leftPanel = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        rows: leftRows,
      });

      const layoutTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        rows: [
          new TableRow({
            cantSplit: true,
            children: [
              new TableCell({
                width: { size: 30, type: WidthType.PERCENTAGE },
                margins: { top: 0, bottom: 0, left: 0, right: 200 } as any,
                verticalAlign: VerticalAlign.TOP,
                borders: {
                  top: { style: BorderStyle.NONE },
                  bottom: { style: BorderStyle.NONE },
                  left: { style: BorderStyle.NONE },
                  right: { style: BorderStyle.NONE },
                },
                children: [leftPanel],
              }),
              new TableCell({
                width: { size: 70, type: WidthType.PERCENTAGE },
                margins: { top: 0, bottom: 0, left: 200, right: 0 } as any,
                verticalAlign: VerticalAlign.TOP,
                borders: {
                  top: { style: BorderStyle.NONE },
                  bottom: { style: BorderStyle.NONE },
                  left: { style: BorderStyle.NONE },
                  right: { style: BorderStyle.NONE },
                },
                children: [mapImgPara],
              }),
            ],
          }),
        ],
      });

      (children as any).push(layoutTable);
    }
  } else {
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 40, after: 120 },
        children: [
          new TextRun({
            text: "Route map not available.",
            size: STYLE.font.cell,
            color: "B42318",
            bold: true,
          }),
        ],
      })
    );
  }

  return {
    properties: sectionPropsA3Landscape(),
    footers: { default: buildFooterTablePages(footerDate ?? new Date()) },
    children,
  };
}

/**
 * GA Drawing can have multiple images. We render ALL images.
 * - If 1 image: single page
 * - If multiple: one image per page (prevents overflow and page-splitting issues)
 */
async function buildGADrawingSections(params: {
  projectName: string;
  gaDrawingBytesList: Uint8Array[];
  footerDate?: string | Date;
}) {
  const { projectName, gaDrawingBytesList, footerDate } = params;

  const list = (gaDrawingBytesList || []).filter((b) => !!b);

  // If nothing, return a single "not available" section.
  if (list.length === 0) {
    const children: Paragraph[] = [];
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [
          new TextRun({
            text: String(projectName || "").toUpperCase(),
            bold: true,
            size: STYLE.font.sectionTitle,
            color: "667085",
          }),
        ],
      })
    );
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: STYLE.spacing.sectionTitle,
        children: [
          new TextRun({
            text: "GA DRAWING FOR 50 FEET TRAILER WITHOUT LOAD:",
            bold: true,
            size: 48,
          }),
        ],
      })
    );
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 40, after: 120 },
        children: [
          new TextRun({
            text: "GA drawing not available.",
            size: STYLE.font.cell,
            color: "B42318",
            bold: true,
          }),
        ],
      })
    );
    return [
      {
        properties: { ...(sectionPropsA3Landscape() as any) } as any,
        footers: { default: buildFooterTablePages(footerDate ?? new Date()) },
        children,
      },
    ];
  }

  // One image per page
  return await Promise.all(list.map(async (bytes, idx) => {
    const children: Paragraph[] = [];

    // Body: keep the drawing image position EXACTLY as before.
    // We add invisible placeholders (same spacing/size as the original title block)
    // so the drawing does NOT jump upward, while the visible titles stay in the header.
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [
          new TextRun({
            text: String(projectName || "").toUpperCase() || " ",
            bold: true,
            size: STYLE.font.sectionTitle,
            color: "FFFFFF",
          }),
        ],
      })
    );
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: STYLE.spacing.sectionTitle,
        children: [
          new TextRun({
            text: "GA DRAWING FOR 50 FEET TRAILER WITHOUT LOAD:",
            bold: true,
            // Match 24pt title height (invisible placeholder)
            size: 48,
            color: "FFFFFF",
          }),
        ],
      })
    );

    children.push(await centeredImageFitTrim(bytes, 1050, 650));

    return {
      // ✅ Keep original GA layout behavior: image/content stays vertically centered on the page.
      // We only move the visible titles to the header (see buildGATitleHeader usage),
      // and keep invisible placeholders in the body to preserve the same drawing position.
      properties: { ...(sectionPropsA3Landscape() as any), verticalAlign: VerticalAlign.CENTER } as any,
      footers: { default: buildFooterTablePages(footerDate ?? new Date()) },
      children,
    };
  }));
}

/** =========================
 * DB Types
 * ========================= */

async function fetchProjectAndLocations4(supabase: any, projectId: string): Promise<{ projectName: string; locs: [string,string,string,string]; debug: string }> {
  if (!projectId) return { projectName: "", locs: ["", "", "", ""] };

  // ✅ "Join" step 1: confirm project row (projects.id) and get its name (or title/project_name)
  const { data: pRow, error: pErr } = await supabase
    .from("projects")
    .select("id, name, title, project_name")
    .eq("id", projectId)
    .maybeSingle();

  if (pErr) throw pErr;
  const pid = pRow?.id || projectId;
  const projectName = String(pRow?.name || pRow?.title || pRow?.project_name || "").trim();

  // ✅ "Join" step 2: fetch locations for that project_id
  const { data: rows, error: lErr } = await supabase
    .from("project_route_page_locations")
    .select("label, sort_order")
    .eq("project_id", pid)
    .order("sort_order", { ascending: true })
    .limit(4);

  if (lErr) throw lErr;

  const labels = (rows || []).map((r: any) => String(r?.label || "").trim());
  const locs: [string, string, string, string] = [
    labels[0] || "",
    labels[1] || "",
    labels[2] || "",
    labels[3] || "",
  ];

  const debug = `projectId=${pid} name=${projectName || '-'} locs=${locs.join(' | ')}`;
  return { projectName, locs, debug };
}

type VehicleMovement = "green" | "yellow" | "red" | "";

type ProjectRow = {
  id: string;
  name?: string | null;
  title?: string | null;
  project_name?: string | null;
};

type ReportRow = {
  id: string;
  project_id: string;
  created_at: string;
  route_id?: string | null;
  category?: string | null;
  description?: string | null;
  difficulty?: string | null;
};

type NormalizedPoint = {
  gps_no: string;
  kms: string;
  ne_coordinate: string;

  details: string;
  location: string;

  description: string;
  movement: VehicleMovement;
  remarks_action: string;

  photo_refs: string[];
  photo_description: string;

  __lat?: number | null;
  __lon?: number | null;
};

function s(v: any) {
  if (v === null || v === undefined) return "";
  return String(v);
}

function projectNameOf(p: ProjectRow | null) {
  return p?.name || p?.title || p?.project_name || "Project";
}

function normalizeMovement(v: any): VehicleMovement {
  const t = String(v ?? "").trim().toLowerCase();
  if (!t) return "";
  if (t === "green") return "green";
  if (t === "yellow" || t === "amber") return "yellow";
  if (t === "red") return "red";
  if (t.includes("red")) return "red";
  if (t.includes("yellow") || t.includes("amber")) return "yellow";
  if (t.includes("green")) return "green";
  return "";
}

/** =========================
 * ✅ PAGE BORDER (tight to edge)
 * ========================= */
function pageBordersTSPL(): any {
  const b = { style: BorderStyle.SINGLE, size: 10, color: PAGE_BORDER_COLOR, space: 0 };
  return {
    borders: {
      pageBorders: {
        top: b,
        left: b,
        bottom: b,
        right: b,
      },
    },
  };
}

/** =========================
 * Text helpers
 * ========================= */
function run(
  text: string,
  opts?: { bold?: boolean; color?: string; size?: number; underline?: boolean }
) {
  return new TextRun({
    text,
    bold: opts?.bold,
    color: opts?.color,
    underline: opts?.underline ? { type: UnderlineType.SINGLE } : undefined,
    size: opts?.size ?? STYLE.font.cell,
  });
}

function paragraphPlain(text: string, align: AlignmentType, spacing?: any, size?: number) {
  return new Paragraph({
    alignment: align,
    spacing: spacing ?? STYLE.spacing.cell,
    children: [new TextRun({ text: text || "—", size: size ?? STYLE.font.cell })],
  });
}

function paragraphFromLine(line: string) {
  const t = (line ?? "").toString().trimEnd();
  const isBullet = t.trim().startsWith("•") || t.trim().startsWith("-") || t.trim().startsWith("• ");
  if (!isBullet) return paragraphPlain(t, AlignmentType.LEFT);

  const normalized = t.trim().startsWith("-") ? `• ${t.trim().slice(1).trim()}` : t.trim();

  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: STYLE.spacing.cell,
    indent: { left: 360, hanging: 180 },
    children: [new TextRun({ text: normalized, size: STYLE.font.cell })],
  });
}

function splitLines(text: string) {
  const lines = (text || "").toString().split("\n").map((x) => x.trimEnd());
  const filtered = lines.filter((x) => x.length > 0);
  return filtered.length ? filtered : ["—"];
}

/** =========================
 * ✅ DETAILS ICONS
 * ========================= */
const DETAILS_ICON_CACHE = new Map<string, Uint8Array>();
const DOCX_ICON_CACHE = new Map<string, Uint8Array | null>();

const CATEGORY_ICON_MAP: Record<string, string> = {
  footpath_bridge: "/images/report-icons/image",
  lt_cable: "/images/report-icons/ca-3",
  ht_cable: "/images/report-icons/ca-3",
  towerline_cable: "/images/report-icons/ca-4",
  diversion: "/images/report-icons/diversion",
  towerline: "/images/report-icons/ca-4",
  underpass: "/images/report-icons/ca-5",
  tree: "/images/report-icons/ca-6",
  bridge: "/images/report-icons/ca-5",
  petrol: "/images/report-icons/ca-15",
  signboard: "/images/report-icons/ca-17",
  electric_sign: "/images/report-icons/ca-17",
  camera_pole: "/images/report-icons/ca-17",
  toll: "/images/report-icons/ca-9",
  junction_left: "/images/report-icons/ca-15",
  bend: "/images/report-icons/ca-13",
  junction_right: "/images/report-icons/ca-16",
  river_bridge: "/images/report-icons/ca-7",
  railway_level_crossing: "/images/report-icons/ca-16",
  gate: "/images/report-icons/ca-11",
};

function buildIconCandidateUrls(src: string): string[] {
  const raw = String(src || "").trim();
  if (!raw) return [];

  const base =
    maybeAbsoluteUrl(raw)

  if (/\.(png|jpg|jpeg)$/i.test(base)) return [base];

  return [
    `${base}.png`,
    `${base}.jpg`,
    `${base}.jpeg`,
    base,
  ];
}

async function getDocxCategoryIcon(kind: string): Promise<Uint8Array | null> {
  if (!kind) return null;
  if (DOCX_ICON_CACHE.has(kind)) return DOCX_ICON_CACHE.get(kind)!;

  const src = CATEGORY_ICON_MAP[kind];

  try {
    if (src) {
      const candidates = buildIconCandidateUrls(src);
      for (const url of candidates) {
        const bytes = await bytesFromUrlForDocx(url);
        if (bytes) {
          DOCX_ICON_CACHE.set(kind, bytes);
          return bytes;
        }
      }
    }

    // Fallback so export still shows an icon even when the external drawing file is missing.
    const fallback = await iconPngBytes(kind, 48);
    DOCX_ICON_CACHE.set(kind, fallback || null);
    return fallback || null;
  } catch {
    const fallback = await iconPngBytes(kind, 48);
    DOCX_ICON_CACHE.set(kind, fallback || null);
    return fallback || null;
  }
}

function normKey(x: string) {
  return String(x || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function detectDetailKind(details: string) {
  const t = normKey(details)
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (t.includes("footpath bridge") || t.includes("foot bridge") || t.includes("pedestrian bridge")) return "footpath_bridge";
  if (t.includes("underpass bridge") || t.includes("underpass")) return "underpass";

  if (t.includes("low tension") || t.includes("lt cable") || t.includes("lt line")) return "lt_cable";
  if (t.includes("high tension") || t.includes("ht cable") || t.includes("ht line")) return "ht_cable";
  if (t.includes("towerline cable") || t.includes("tower line cable")) return "towerline_cable";
  if (t === "towerline" || t.includes("towerline") || t.includes("tower line")) return "towerline";

  if (t.includes("junction left") || t.includes("left junction") || t.includes("turn left")) return "junction_left";
  if (t.includes("junction right") || t.includes("right junction") || t.includes("turn right")) return "junction_right";
  if (t === "bend" || t.includes("bend") || t.includes("curve")) return "bend";
  if (t.includes("take diversion") || t == "diversion" || t.includes("diversion")) return "diversion";

  if (t.includes("tree branches") || t.includes("tree branch") || t == "tree" || t.includes("branches")) return "tree";
  if (t.includes("petrol bunk") || t.includes("petrol pump") || t.includes("fuel station") || t.includes("fuel bunk")) return "petrol";
  if (t.includes("electric sign board") || t.includes("electric signboard") || t.includes("illuminated sign")) return "electric_sign";
  if (t.includes("signboard") || t.includes("sign board") || t.includes("road sign")) return "signboard";
  if (t.includes("camera pole") || t.includes("camera") || t.includes("cctv pole") || t.includes("surveillance pole")) return "camera_pole";
  if (t.includes("toll plaza") || t == "toll" || t.includes("toll")) return "toll";

  if (t === "bridge" || t.includes(" bridge")) return "bridge";

  return "";
}

async function iconPngBytes(kind: string, sizePx = 34): Promise<Uint8Array | null> {
  const key = `${kind}:${sizePx}`;
  if (DETAILS_ICON_CACHE.has(key)) return DETAILS_ICON_CACHE.get(key)!;
  try {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${sizePx}" height="${sizePx}" viewBox="0 0 ${sizePx} ${sizePx}"><rect x="2" y="2" width="${sizePx-4}" height="${sizePx-4}" rx="4" ry="4" fill="white" stroke="#111" stroke-width="2"/><text x="50%" y="54%" text-anchor="middle" dominant-baseline="middle" font-family="Arial" font-size="${Math.max(9, Math.floor(sizePx*0.24))}" font-weight="700" fill="#111">${String(kind || "OBS").replace(/[^A-Za-z]/g, "").slice(0,4).toUpperCase() || "OBS"}</text></svg>`;
    const out = await sharp(Buffer.from(svg)).png().toBuffer();
    const bytes = new Uint8Array(out);
    DETAILS_ICON_CACHE.set(key, bytes);
    return bytes;
  } catch {
    return null;
  }
}

/** =========================
 * ✅ Watermark (diagonal)
 * ========================= */
async function watermarkPngBytesDiagonal(text: string) {
  try {
    const safe = String(text || "CONFIDENTIAL").replace(/[<&>]/g, " ");
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900"><g transform="translate(800,450) rotate(-30)"><text x="0" y="0" text-anchor="middle" dominant-baseline="middle" font-family="Arial" font-size="160" font-weight="700" fill="rgba(120,120,120,0.12)">${safe}</text></g></svg>`;
    return new Uint8Array(await sharp(Buffer.from(svg)).png().toBuffer());
  } catch {
    return new Uint8Array();
  }
}

async function buildHeaderWithDiagonalWatermark() {
  // Watermark removed per requirement
  return new Header({ children: [] });
}


/** =========================
 * Footers
 * ========================= */
function formatDDMMYYYY(input?: string | Date) {
  const d = input ? new Date(input) : new Date();
  if (Number.isNaN(d.getTime())) return formatDDMMYYYY(new Date());
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = String(d.getFullYear());
  return `${dd}.${mm}.${yyyy}`;
}

function formatDDMMYYYY_DASH(input?: string | Date) {
  const d = input ? new Date(input) : new Date();
  if (Number.isNaN(d.getTime())) return formatDDMMYYYY_DASH(new Date());
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = String(d.getFullYear());
  return `${dd}-${mm}-${yyyy}`;
}

function buildFooterTablePages(date?: string | Date) {
  const dateStr = formatDDMMYYYY(date);
  const siteText = "raceinnovations.in";
  const siteUrl = "https://raceinnovations.in";

  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

  const row = new TableRow({
    children: [
      new TableCell({
        width: { size: 33, type: WidthType.PERCENTAGE },
        borders: { top: none, left: none, right: none, bottom: none },
        children: [
          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: STYLE.spacing.none,
            children: [new TextRun({ text: `Date : ${dateStr}`, size: STYLE.font.meta })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 34, type: WidthType.PERCENTAGE },
        borders: { top: none, left: none, right: none, bottom: none },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: STYLE.spacing.none,
            children: [
              new TextRun({ text: "CONFIDENTIAL ", bold: true, size: STYLE.font.meta }),
              new ExternalHyperlink({
                link: siteUrl,
                children: [
                  new TextRun({
                    text: siteText,
                    size: STYLE.font.meta,
                    color: "0000FF",
                    underline: { type: UnderlineType.SINGLE },
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new TableCell({
        width: { size: 33, type: WidthType.PERCENTAGE },
        borders: { top: none, left: none, right: none, bottom: none },
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: STYLE.spacing.none,
            children: [new TextRun({ children: ["PAGE NO. ", PageNumber.CURRENT], size: STYLE.font.meta })],
          }),
        ],
      }),
    ],
  });

  return new Footer({
    children: [
      new Table({
        layout: TableLayoutType.FIXED,
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [row],
      }),
    ],
  });
}

function buildCoverFooter(opts: CoverOptions) {
  const leftText = opts.footerLeftText ?? "Report by RACE Innovations Pvt ltd";
  const email = opts.footerEmail ?? "kh@raceinnovations.in";
  const website = opts.footerWebsite ?? "https://raceinnovations.in/";
  const datedLabel = opts.datedLabel ?? "Dated";
  const dated = formatDDMMYYYY_DASH(opts.date);

  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: STYLE.spacing.none,
        tabStops: [{ type: "right", position: 12500 }],
        children: [
          run(`${leftText} | `, { size: STYLE.font.meta }),
          run(`email at `, { size: STYLE.font.meta }),
          run(email, { size: STYLE.font.meta, color: "0563C1", underline: true }),
          run(`  |  `, { size: STYLE.font.meta }),
          new ExternalHyperlink({
            link: website.startsWith("http") ? website : `https://${website.replace(/^\/+/, "")}`,
            children: [
              new TextRun({
                text: website,
                size: STYLE.font.meta,
                color: "0563C1",
                underline: { type: UnderlineType.SINGLE },
              }),
            ],
          }),
          run(`\t${datedLabel} ${dated}`, { size: STYLE.font.meta, color: "808080" }),
        ],
      }),
    ],
  });
}


// Logo-only paragraph used in headers/cover. Keep height modest to avoid a "stretched" look.
async function coverLogoOnly(logoUrl?: string, w = 390, h = 68) {
  // NOTE: In Next.js, assets under /public are served from the site root.
  // Many projects store this logo at /public/images/logo_v2.png (NOT /public/logo_v2.png).
  // To prevent "logo not showing" due to an incorrect path, try a small set of candidates.
  const raw = String(logoUrl || "").trim();
  const candidates = [
    raw,
    raw || "/images/logo_v2.png",
    raw || "/logo_v2.png",
  ].filter(Boolean) as string[];

  const toAbs = (u0: string) =>
    maybeAbsoluteUrl(u0);

  let bytes: Uint8Array | null = null;
  for (const c of candidates) {
    bytes = await bytesFromUrlForDocx(toAbs(c));
    if (bytes) break;
  }
  if (!bytes) {
    // Logo requested, but not available. Do NOT render fallback text.
    return new Paragraph({ alignment: AlignmentType.LEFT, spacing: STYLE.spacing.none, text: "" });
  }
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 0 },
    children: [new ImageRun({ data: bytes, transformation: { width: w, height: h } })],
  });
}


async function buildLogoOnlyHeader(logoUrl?: string, w = 330, h = 58) {
  const raw = String(logoUrl || "").trim();
  const candidates = [raw, raw || "/images/logo_v2.png", raw || "/logo_v2.png"].filter(Boolean) as string[];
  const toAbs = (u0: string) =>
    maybeAbsoluteUrl(u0);

  let bytes: Uint8Array | null = null;
  for (const c of candidates) {
    bytes = await bytesFromUrlForDocx(toAbs(c));
    if (bytes) break;
  }
  if (!bytes) return new Header({ children: [] });

  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 0 },
        children: [new ImageRun({ data: bytes, transformation: { width: w, height: h } })],
      }),
    ],
  });
}

// ✅ GA header: logo + titles at top (keeps page body content unchanged)
async function buildGATitleHeader(params: {
  logoUrl?: string;
  projectName: string;
  includeGATitle?: boolean;
  logoW?: number;
  logoH?: number;
}) {
  const { logoUrl, projectName, includeGATitle = true, logoW = 220, logoH = 38 } = params;

  const raw = String(logoUrl || "").trim();
  const candidates = [raw, raw || "/images/logo_v2.png", raw || "/logo_v2.png"].filter(Boolean) as string[];
  const toAbs = (u0: string) =>
    maybeAbsoluteUrl(u0);

  let bytes: Uint8Array | null = null;
  for (const c of candidates) {
    bytes = await bytesFromUrlForDocx(toAbs(c));
    if (bytes) break;
  }

  const children: Paragraph[] = [];

  if (bytes) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 20 },
        children: [new ImageRun({ data: bytes, transformation: { width: logoW, height: logoH } })],
      })
    );
  }

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: includeGATitle ? 40 : 0 },
      children: [
        new TextRun({
          text: String(projectName || "").toUpperCase(),
          bold: true,
          size: STYLE.font.sectionTitle,
          color: "667085",
        }),
      ],
    })
  );

  if (includeGATitle) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 0 },
        children: [
          new TextRun({
            text: "GA DRAWING FOR 50 FEET TRAILER WITHOUT LOAD:",
            bold: true,
            // User-required: page title = 24pt => 48 half-points
            size: 48,
          }),
        ],
      })
    );
  }

  return new Header({ children });
}


/** =========================
 * Cover helpers
 * ========================= */
async function coverTopRowLogo(cover: CoverOptions, centerText: string) {
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

  // Left: Logo (image)
  const logoPara = await coverLogoOnly(cover.logoUrl, cover.logoWidth ?? 320, cover.logoHeight ?? 55);

  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        height: { value: 900, rule: HeightRule.ATLEAST },
        children: [
          new TableCell({
            width: { size: 33, type: WidthType.PERCENTAGE },
            borders: { top: none, bottom: none, left: none, right: none },
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            verticalAlign: VerticalAlign.TOP,
            children: [logoPara],
          }),
          new TableCell({
            width: { size: 34, type: WidthType.PERCENTAGE },
            borders: { top: none, bottom: none, left: none, right: none },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: STYLE.spacing.none,
                children: [new TextRun({ text: String(centerText || "").toUpperCase(), size: 32, color: "1F4E79" })],
              }),
            ],
          }),
          // Right: EMPTY (CI CHANNEL'S INDIA removed)
          new TableCell({
            width: { size: 33, type: WidthType.PERCENTAGE },
            borders: { top: none, bottom: none, left: none, right: none },
            children: [new Paragraph({ spacing: STYLE.spacing.none, text: "" })],
          }),
        ],
      }),
    ],
  });
}

async function buildCoverHeader(cover: CoverOptions, centerText: string) {
  const table = await coverTopRowLogo(cover, centerText);
  return new Header({ children: [table] });
}


function coverLine(color = "1F4E79") {
  return new Paragraph({
    spacing: { before: 180, after: 180 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color } },
  } as any);
}

function coverTitleProject(projectName: string) {
  const t = String(projectName || "").trim() || "PROJECT";
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    // Keep the cover title visually tight (reference uses compact vertical rhythm)
    spacing: { before: 0, after: 120, line: 520 },
    // User-required: Center title (CBE) = 36pt => 72 half-points
    children: [new TextRun({ text: t.toUpperCase(), bold: true, size: 72, color: "1F3A5F" })],
  });
}

function coverRecommendationBox(text: string) {
  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 62, type: WidthType.PERCENTAGE },
    alignment: AlignmentType.CENTER as any,
    indent: { size: 0, type: WidthType.DXA } as any,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { type: ShadingType.CLEAR, fill: "FFFFFF" },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 10, color: "1F4E79" },
              bottom: { style: BorderStyle.SINGLE, size: 10, color: "1F4E79" },
              left: { style: BorderStyle.SINGLE, size: 10, color: "1F4E79" },
              right: { style: BorderStyle.SINGLE, size: 10, color: "1F4E79" },
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 220, after: 220 },
                // User-required: SURVEY REPORT = 24pt => 48 half-points
                children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 48, color: "1F3A5F" })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

/** =========================
 * Table header helpers
 * ========================= */
function headerCell(text: string, span: number) {
  return new TableCell({
    columnSpan: span,
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: HEADER_FILL },
    borders: CELL_BORDERS,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: STYLE.spacing.none,
        children: [run(text, { bold: true, color: "FFFFFF", size: STYLE.font.header })],
      }),
    ],
  });
}

function makeHeaderRow() {
  return new TableRow({
    children: [
      headerCell("GPS NO", 1),
      headerCell("KMS", 1),
      headerCell("NE CO-ORDINATE", 1),
      headerCell("DETAILS", 2),
      headerCell("LOCATION", 2),
      headerCell("DESCRIPTION", 1),
      headerCell("STATUS", 1),
      headerCell("PHOTO", 3),
    ],
  });
}

function textCell(text: string, span: number, align: AlignmentType, vAlign: VerticalAlign) {
  const lines = splitLines(text);
  const paras =
    align === AlignmentType.LEFT ? lines.map(paragraphFromLine) : lines.map((ln) => paragraphPlain(ln, align, STYLE.spacing.cell));

  return new TableCell({
    columnSpan: span,
    verticalAlign: vAlign,
    borders: CELL_BORDERS,
    margins: STYLE.cellMargins,
    children: paras,
  });
}

/** =========================
 * DETAILS cell: icon + text
 * ========================= */
async function detailsCellWithIcon(detailsText: string, span: number, vAlign: VerticalAlign) {
  const kind = detectDetailKind(detailsText);
  const lines = splitLines(detailsText);
  const firstLine = lines[0] ?? "—";
  const rest = lines.slice(1);

  const iconBytes = kind ? await getDocxCategoryIcon(kind) : null;

  const firstPara = new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: STYLE.spacing.cell,
    children: [
      ...(iconBytes
        ? [new ImageRun({ data: iconBytes, transformation: { width: 42, height: 42 } }), new TextRun({ text: "  ", size: STYLE.font.cell })]
        : []),
      new TextRun({ text: firstLine, size: STYLE.font.cell }),
    ],
  });

  const restParas = rest.map((ln) => paragraphFromLine(ln));

  return new TableCell({
    columnSpan: span,
    verticalAlign: vAlign,
    borders: CELL_BORDERS,
    margins: STYLE.cellMargins,
    children: [firstPara, ...restParas],
  });
}

/** =========================
 * VEHICLE MOVEMENT square
 * ========================= */
async function squarePngBytes(colorHex: string, sizePx = 26): Promise<Uint8Array> {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${sizePx}" height="${sizePx}"><rect x="1" y="1" width="${sizePx-2}" height="${sizePx-2}" fill="${colorHex}" stroke="#111" stroke-width="2"/></svg>`;
  return new Uint8Array(await sharp(Buffer.from(svg)).png().toBuffer());
}

const MOVEMENT_SQUARE_CACHE = new Map<string, Uint8Array>();

async function movementCell(movement: string) {
  const m = normalizeMovement(movement);
  const color =
    m === "red" ? "#FF0000" : m === "yellow" ? "#FFC000" : m === "green" ? "#00B050" : "#FFFFFF";

  const box = 24;
  const key = `${color}:${box}`;

  let bytes = MOVEMENT_SQUARE_CACHE.get(key);
  if (!bytes) {
    bytes = await squarePngBytes(color, box);
    MOVEMENT_SQUARE_CACHE.set(key, bytes);
  }

  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    borders: CELL_BORDERS,
    margins: { top: 0, bottom: 0, left: 0, right: 0 } as any,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: STYLE.spacing.none,
        children: [new ImageRun({ data: bytes, transformation: { width: box, height: box } })],
      }),
    ],
  });
}

/** =========================
 * Photo detection helpers
 * ========================= */
function looksLikeImageRef(str: string) {
  const t = str.trim();
  if (!t) return false;
  if (/^https?:\/\//i.test(t)) return true;
  if (/\.(jpe?g|png|webp|gif|bmp|heic)(\?.*)?$/i.test(t)) return true;
  if (t.includes("storage/v1/object")) return true;
  if (t.includes("/") && t.length > 8) return true;
  return false;
}

function collectImageStrings(value: any, out: string[] = [], seen = new Set<any>(), depth = 0) {
  if (depth > 6) return out;
  if (value === null || value === undefined) return out;

  if (typeof value === "object") {
    if (seen.has(value)) return out;
    seen.add(value);
  }

  if (typeof value === "string") {
    const t = value.trim();
    if (!t) return out;

    if ((t.startsWith("[") && t.endsWith("]")) || (t.startsWith("{") && t.endsWith("}"))) {
      try {
        return collectImageStrings(JSON.parse(t), out, seen, depth + 1);
      } catch {}
    }

    if (t.includes(",")) {
      t.split(",").forEach((x) => collectImageStrings(x, out, seen, depth + 1));
      return out;
    }

    if (looksLikeImageRef(t)) out.push(t);
    return out;
  }

  if (Array.isArray(value)) {
    for (const v of value) collectImageStrings(v, out, seen, depth + 1);
    return out;
  }

  if (typeof value === "object") {
    const maybeUrl = (value as any).url ?? (value as any).path ?? (value as any).signedUrl;
    if (typeof maybeUrl === "string") collectImageStrings(maybeUrl, out, seen, depth + 1);
    for (const v of Object.values(value)) collectImageStrings(v, out, seen, depth + 1);
    return out;
  }

  return out;
}

/** =========================
 * Supabase storage resolve
 * ========================= */
const BUCKET_CANDIDATES = [
  "ga-drawings",
  "ga_drawings",
  "route-maps",
  "route_maps",
  "report-photos",
  "report_photos",
  "report-images",
  "report_images",
  "route-photos",
  "route-images",
  "route_images",
  "route_photos",
  "project-photos",
  "project_photos",
  "project-images",
  "project_images",
  "media",
  "uploads",
  "files",
  "attachments",
  "images",
  "photos",
  "report-media",
  "report_media",
  "route-media",
  "route_media",
];

function isAbsoluteUrl(u: string) {
  return /^https?:\/\//i.test(u) || u.startsWith("data:");
}

function cleanPath(p: string) {
  return p.replace(/^\/+/, "");
}

async function blobToBytes(b: Blob) {
  return new Uint8Array(await b.arrayBuffer());
}


const DOCX_OPTIMIZED_PHOTO_CACHE = new Map<string, Uint8Array | null>();

async function optimizeImageBytesForDocx(
  bytes: Uint8Array,
  opts: { maxWidth: number; maxHeight: number; quality?: number }
): Promise<Uint8Array> {
  try {
    const img = sharp(Buffer.from(bytes), { density: 144, pages: 1 });
    const meta = await img.metadata();
    const srcW = Math.max(1, meta.width || opts.maxWidth);
    const srcH = Math.max(1, meta.height || opts.maxHeight);
    const scale = Math.min(1, opts.maxWidth / srcW, opts.maxHeight / srcH);
    const outW = Math.max(1, Math.round(srcW * scale));
    const outH = Math.max(1, Math.round(srcH * scale));
    const out = await img.resize(outW, outH, { fit: 'inside', withoutEnlargement: true }).jpeg({ quality: Math.round((opts.quality ?? 0.76) * 100) }).toBuffer();
    return new Uint8Array(out);
  } catch {
    return bytes;
  }
}

async function resolvePhotoBytesForDocx(
  supabase: any,
  ref: string,
  variant: "cell" | "page" = "cell"
): Promise<Uint8Array | null> {
  const key = `${variant}:${ref}`;
  if (DOCX_OPTIMIZED_PHOTO_CACHE.has(key)) {
    return DOCX_OPTIMIZED_PHOTO_CACHE.get(key)!;
  }

  const raw = await resolvePhotoBytes(supabase, ref);
  if (!raw) {
    DOCX_OPTIMIZED_PHOTO_CACHE.set(key, null);
    return null;
  }

  const optimized = await optimizeImageBytesForDocx(
    raw,
    variant === "page"
      ? { maxWidth: 1800, maxHeight: 1200, quality: 0.8 }
      : { maxWidth: 1200, maxHeight: 900, quality: 0.72 }
  );

  DOCX_OPTIMIZED_PHOTO_CACHE.set(key, optimized);
  return optimized;
}

const PHOTO_BYTES_CACHE = new Map<string, Uint8Array | null>();
let BUCKET_NAMES_CACHE: string[] | null = null;
let BUCKET_NAMES_PROMISE: Promise<string[]> | null = null;

const DEFAULT_PHOTO_TIMEOUT_MS = 12_000;
const DEFAULT_STORAGE_TIMEOUT_MS = 12_000;
const DEFAULT_BUCKETS_TIMEOUT_MS = 10_000;

function withTimeout<T>(p: Promise<T>, ms: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const t = setTimeout(() => reject(new Error("timeout")), ms);
    p.then((v) => {
      clearTimeout(t);
      resolve(v);
    }).catch((e) => {
      clearTimeout(t);
      reject(e);
    });
  });
}

async function safeTimeout<T>(p: Promise<T>, ms: number): Promise<T | null> {
  try {
    return await withTimeout(p, ms);
  } catch {
    return null;
  }
}

async function getBucketNamesOnce(supabase: any): Promise<string[]> {
  if (BUCKET_NAMES_CACHE) return BUCKET_NAMES_CACHE;

  if (!BUCKET_NAMES_PROMISE) {
    BUCKET_NAMES_PROMISE = (async () => {
      const res: any = await safeTimeout(supabase.storage.listBuckets(), DEFAULT_BUCKETS_TIMEOUT_MS);
      const names = Array.isArray(res?.data) ? res.data.map((b: any) => b?.name).filter(Boolean) : [];
      return Array.from(new Set(names));
    })();
  }

  const names = (await BUCKET_NAMES_PROMISE) || [];
  BUCKET_NAMES_CACHE = names;
  return BUCKET_NAMES_CACHE;
}

async function blobToPngBytes(blob: Blob): Promise<Uint8Array | null> {
  try {
    const buf = Buffer.from(await blob.arrayBuffer());
    const out = await sharp(buf, { density: 144, pages: 1 }).png().toBuffer();
    return new Uint8Array(out);
  } catch {
    return null;
  }
}

const PDF_JS_PROMISE_KEY = "__docxPdfJsPromise";

async function getPdfJsForDocx() {
  return null as any;
}

async function pdfBlobToPngBytes(blob: Blob): Promise<Uint8Array | null> {
  try {
    const buf = Buffer.from(await blob.arrayBuffer());
    const out = await sharp(buf, { density: 144, pages: 1 }).png().toBuffer();
    return new Uint8Array(out);
  } catch {
    return null;
  }
}

async function fetchBytes(url: string, timeoutMs = DEFAULT_PHOTO_TIMEOUT_MS) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const localBytes = url.startsWith("/") ? await readLocalPublicFileBytes(url) : null;
    if (localBytes) return localBytes;
    const res = await fetch(maybeAbsoluteUrl(url), { signal: controller.signal, cache: "no-store" });
    if (!res.ok) throw new Error(`Photo fetch failed: ${res.status}`);

    const blob = await res.blob();
    const type = String(blob.type || "").toLowerCase();
    const looksPdf = type === "application/pdf" || /\.pdf($|[?#])/i.test(url);

    if (looksPdf) {
      const pdfBytes = await pdfBlobToPngBytes(blob);
      if (pdfBytes) return pdfBytes;
    }

    // DOCX image rendering is most reliable with PNG/JPEG.
    // Convert GIF/WEBP/SVG/AVIF/BMP and other browser-decodable formats to PNG first.
    if (type && type !== "image/png" && type !== "image/jpeg" && type !== "image/jpg") {
      const pngBytes = await blobToPngBytes(blob);
      if (pngBytes) return pngBytes;
    }

    return new Uint8Array(await blob.arrayBuffer());
  } finally {
    clearTimeout(t);
  }
}

async function tryDownloadThenSignedThenPublic(supabase: any, bucket: string, path: string): Promise<Uint8Array | null> {
  const dl: any = await safeTimeout(supabase.storage.from(bucket).download(path), DEFAULT_STORAGE_TIMEOUT_MS);
  if (dl && !dl.error && dl.data) {
    try {
      const blob = dl.data as Blob;
      const type = String(blob?.type || "").toLowerCase();
      const looksPdf = type === "application/pdf" || /\.pdf$/i.test(path);

      if (looksPdf) {
        const pdfBytes = await pdfBlobToPngBytes(blob);
        if (pdfBytes) return pdfBytes;
      }

      if (type && type !== "image/png" && type !== "image/jpeg" && type !== "image/jpg") {
        const pngBytes = await blobToPngBytes(blob);
        if (pngBytes) return pngBytes;
      }

      return await blobToBytes(blob);
    } catch {}
  }

  const signed: any = await safeTimeout(
    supabase.storage.from(bucket).createSignedUrl(path, 60 * 10),
    DEFAULT_STORAGE_TIMEOUT_MS
  );
  if (signed && !signed.error && signed.data?.signedUrl) {
    try {
      return await fetchBytes(signed.data.signedUrl);
    } catch {}
  }

  try {
    const pub = supabase.storage.from(bucket).getPublicUrl(path);
    const publicUrl = pub?.data?.publicUrl;
    if (publicUrl) return await fetchBytes(publicUrl);
  } catch {}

  return null;
}

function extractBucketAndPathFromStorageUrl(url: string): { bucket: string; path: string } | null {
  try {
    const u = new URL(url);

    const wrapped = u.searchParams.get("url");
    if (wrapped && wrapped !== url) {
      return extractBucketAndPathFromStorageUrl(decodeURIComponent(wrapped));
    }

    const pathname = decodeURIComponent(u.pathname || "");

    const patterns = [
      "/storage/v1/object/public/",
      "/storage/v1/object/sign/",
      "/storage/v1/object/authenticated/",
      "/storage/v1/render/image/public/",
    ];

    for (const marker of patterns) {
      const idx = pathname.indexOf(marker);
      if (idx === -1) continue;

      const tail = pathname.slice(idx + marker.length);
      const parts = tail.split("/").filter(Boolean);
      if (parts.length < 2) continue;

      const bucket = parts[0];
      const path = parts.slice(1).join("/");
      if (!bucket || !path) continue;

      return { bucket, path: cleanPath(path) };
    }

    return null;
  } catch {
    return null;
  }
}

async function resolvePhotoBytes(supabase: any, ref: string): Promise<Uint8Array | null> {
  const raw = (ref || "").trim();
  if (!raw) return null;
  if (PHOTO_BYTES_CACHE.has(raw)) return PHOTO_BYTES_CACHE.get(raw) as any;

  const out = await (async (): Promise<Uint8Array | null> => {
    if (isAbsoluteUrl(raw)) {
      const parsed = extractBucketAndPathFromStorageUrl(raw);
      if (parsed) {
        const viaApi = await tryDownloadThenSignedThenPublic(supabase, parsed.bucket, parsed.path);
        if (viaApi) return viaApi;
      }

      try {
        const u = new URL(raw);
        const wrapped = u.searchParams.get("url");
        if (wrapped && wrapped !== raw) {
          const inner = decodeURIComponent(wrapped);
          const parsedInner = extractBucketAndPathFromStorageUrl(inner);
          if (parsedInner) {
            const viaInnerApi = await tryDownloadThenSignedThenPublic(supabase, parsedInner.bucket, parsedInner.path);
            if (viaInnerApi) return viaInnerApi;
          }
          try {
            return await fetchBytes(inner);
          } catch {}
        }
      } catch {}

      try {
        return await fetchBytes(raw);
      } catch {
        return null;
      }
    }

    const cleaned = cleanPath(decodeURIComponent(raw));
    const parts = cleaned.split("/");
    const first = parts[0] || "";

    let bucketsToTry = [...BUCKET_CANDIDATES];
    let pathToUse = cleaned;

    if (bucketsToTry.includes(first) && parts.length > 1) {
      bucketsToTry = [first, ...bucketsToTry.filter((b) => b !== first)];
      pathToUse = parts.slice(1).join("/");
    }

    const dynamic = await getBucketNamesOnce(supabase);
    if (dynamic.length) bucketsToTry = Array.from(new Set([...dynamic, ...bucketsToTry]));

    for (const bucket of bucketsToTry) {
      const bytes = await tryDownloadThenSignedThenPublic(supabase, bucket, pathToUse);
      if (bytes) return bytes;
    }

    return null;
  })().catch(() => null);

  PHOTO_BYTES_CACHE.set(raw, out);
  return out;
}

/** =========================
 * PHOTO cell: adjusted sizing so the image is fully visible in the widened PHOTO area
 * ========================= */
async function photoCell(supabase: any, refs: string[], includePhotos: boolean, caption?: string) {
  const list = (refs || []).filter(Boolean).slice(0, 3);

  if (!includePhotos || list.length === 0) {
    const cap = (caption || "").trim();
    return new TableCell({
      columnSpan: 3,
      verticalAlign: VerticalAlign.TOP,
      borders: PHOTO_CELL_BORDERS,
      margins: { top: 40, bottom: 40, left: 80, right: 80 } as any,
      children: cap
        ? [
            new Paragraph({
              alignment: AlignmentType.LEFT,
              spacing: STYLE.spacing.cell,
              children: [new TextRun({ text: cap, size: STYLE.font.cellSmall })],
            }),
          ]
        : [new Paragraph({ spacing: STYLE.spacing.none, text: "" })],
    });
  }

  const multi = list.length > 1;
  const imgW = multi ? STYLE.photo.multi.w : STYLE.photo.single.w;
  const imgH = multi ? STYLE.photo.multi.h : STYLE.photo.single.h;

  const bytesList = await Promise.all(list.map((r) => resolvePhotoBytesForDocx(supabase, r, "cell")));

  const paras: Paragraph[] = [];
  for (let i = 0; i < bytesList.length; i++) {
    const bytes = bytesList[i];
    if (!bytes) continue;
    paras.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { ...STYLE.spacing.none, after: i === bytesList.length - 1 ? 0 : 40 },
        children: [new ImageRun({ data: bytes, transformation: { width: imgW, height: imgH } })],
      })
    );
  }

  if (paras.length === 0 && DEBUG_PHOTOS) {
    throw new Error(`Photo refs detected but could not be resolved.\nFirst ref: ${list[0]}`);
  }

  const cap2 = (caption || "").trim();
  if (cap2) {
    paras.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 80, after: 0, line: 276 },
        children: [new TextRun({ text: cap2, size: STYLE.font.cellSmall })],
      })
    );
  }

  return new TableCell({
    columnSpan: 3,
    verticalAlign: VerticalAlign.TOP,
    borders: PHOTO_CELL_BORDERS,
    margins: { top: 40, bottom: 40, left: 80, right: 80 } as any,
    children: paras.length ? paras : [new Paragraph({ spacing: STYLE.spacing.none, text: "" })],
  });
}

/** =========================
 * Normalize Point
 * ========================= */
function normalizePoint(raw: any): NormalizedPoint {
  const gpsCandidate = s(
    raw.gps_no ??
      raw.gps ??
      raw.no ??
      raw.sno ??
      raw.sl_no ??
      raw.serial ??
      raw.seq ??
      raw.sequence ??
      raw.point_no ??
      raw.point_index ??
      raw.idx ??
      raw.index ??
      ""
  );

  const kmsCandidate = s(
    raw.kms ??
      raw.km ??
      raw.distance ??
      raw.dist ??
      raw.chainage ??
      raw.ch ??
      raw.kilometer ??
      raw.kilometre ??
      raw.route_km ??
      raw.km_value ??
      ""
  );

  const rawNe =
    typeof raw.ne_coordinate === "string" && raw.ne_coordinate.trim()
      ? raw.ne_coordinate.trim()
      : typeof raw.coordinate === "string" && raw.coordinate.trim()
        ? raw.coordinate.trim()
        : "";

  let lat: number | null = null;
  let lon: number | null = null;
  let ne_coordinate = "";

  if (rawNe) {
    ne_coordinate = rawNe;
    const parsed = parseNEToDecimal(rawNe);
    if (parsed) {
      lat = parsed.lat;
      lon = parsed.lon;
    }
  } else {
    const latRaw = raw.loc_lat ?? raw.lat ?? raw.latitude ?? raw.north ?? raw.n;
    const lonRaw = raw.loc_lon ?? raw.lon ?? raw.lng ?? raw.longitude ?? raw.east ?? raw.e;

    const latNum = latRaw != null ? Number(latRaw) : null;
    const lonNum = lonRaw != null ? Number(lonRaw) : null;

    if (
      latNum != null &&
      lonNum != null &&
      !Number.isNaN(latNum) &&
      !Number.isNaN(lonNum)
    ) {
      lat = latNum;
      lon = lonNum;
      ne_coordinate = formatNEFromLatLon(latNum, lonNum);
    }
  }

  const details = s(
    raw.details ?? raw.remarks ?? raw.note ?? raw.__report_category ?? raw.category ?? ""
  );
  const location = s(raw.exact_location ?? raw.location_name ?? raw.location_label ?? raw.location ?? raw.place ?? raw.area ?? raw.city ?? raw.village ?? "");
  const photo_refs = Array.from(new Set(collectImageStrings(raw)));

  const photo_description = s(
    raw.photo_description ??
      raw.photo_desc ??
      raw.image_description ??
      raw.__report_description ??
      raw.description ??
      raw.desc ??
      ""
  );

  let description = s(raw.description ?? raw.desc ?? raw.__report_description ?? "").trim();

  if (description && details && description.toLowerCase() === details.trim().toLowerCase()) {
    description = "";
  }

  const movement = normalizeMovement(
    raw.difficulty ??
      raw.vehicle_movement ??
      raw.movement ??
      raw.status ??
      raw.__report_difficulty ??
      ""
  );

  const remarks_action = s(
    raw.remarks_action ??
      raw.action ??
      raw.actions ??
      raw.__report_remarks_action ??
      raw.__report_difficulty ??
      ""
  ).trim();

  return {
    gps_no: gpsCandidate,
    kms: kmsCandidate,
    ne_coordinate,
    details,
    location,
    description: description || "—",
    movement,
    remarks_action,
    photo_refs,
    photo_description,
    __lat: lat,
    __lon: lon,
  };
}

/** =========================
 * KMS compute
 * ========================= */
function enrichPointsAlways(points: any[]): NormalizedPoint[] {
  const norm = points.map((p) => normalizePoint(p));

  for (let i = 0; i < norm.length; i++) {
    if (!norm[i].gps_no) norm[i].gps_no = String(i + 1);
  }

  let cum = 0;
  for (let i = 0; i < norm.length; i++) {
    const cur = norm[i];
    const prev = i > 0 ? norm[i - 1] : null;

    if (!cur.kms) {
      if (
        prev?.__lat != null &&
        prev?.__lon != null &&
        cur.__lat != null &&
        cur.__lon != null &&
        !Number.isNaN(prev.__lat) &&
        !Number.isNaN(prev.__lon) &&
        !Number.isNaN(cur.__lat) &&
        !Number.isNaN(cur.__lon)
      ) {
        cum += haversineKm(prev.__lat, prev.__lon, cur.__lat, cur.__lon);
        cur.kms = cum.toFixed(4);
      } else {
        cur.kms = i === 0 ? "0.0000" : "—";
      }
    }
  }

  return norm;
}

/** =========================
 * Extra photos (report_photos)
 * ========================= */
async function getExtraPhotosForReport(supabase: any, reportId: string) {
  const { data, error } = await supabase
    .from("report_photos")
    .select("url, created_at")
    .eq("report_id", reportId)
    .order("created_at", { ascending: true })
    .limit(300);

  if (error || !Array.isArray(data)) return [];
  const refs = data.map((r: any) => String(r?.url || "").trim()).filter(Boolean);
  return Array.from(new Set(refs));
}

function applyExtraPhotos(points: NormalizedPoint[], extraRefs: string[]) {
  if (!extraRefs.length) return points;

  let j = 0;
  for (let i = 0; i < points.length && j < extraRefs.length; i++) {
    const has = Array.isArray(points[i].photo_refs) && points[i].photo_refs.length > 0;
    if (!has) {
      points[i].photo_refs = [extraRefs[j]];
      j++;
    }
  }

  if (j < extraRefs.length && points.length) {
    const last = points[points.length - 1];
    const merged = Array.from(new Set([...(last.photo_refs || []), ...extraRefs.slice(j)]));
    last.photo_refs = merged.slice(0, 3);
  }

  return points;
}

/** =========================
 * Points loader
 * ========================= */
const TABLE_CANDIDATES = [
  "report_path_points",
  "route_points",
  "route_point",
  "route_locations",
  "route_location",
  "route_logs",
  "gps_logs",
  "gps_points",
  "location_logs",
  "locations",
  "location_points",
  "track_points",
  "tracking_points",
  "report_points",
  "report_point",
  "report_items",
  "report_entries",
  "report_details",
  "points",
];

const FK_CANDIDATES = [
  "report_id",
  "reportid",
  "reportId",
  "reports_id",
  "parent_report_id",
  "report_ref",
  "report_uuid",
  "route_id",
  "routeid",
  "routeId",
  "project_id",
  "projectid",
  "projectId",
];

async function getPointsForReport(supabase: any, reportId: string) {
  const { data: report, error: rErr } = await supabase.from("reports").select("*").eq("id", reportId).single();
  if (rErr) throw rErr;

  const routeId = report?.route_id ?? report?.routeId ?? null;
  const projectId = report?.project_id ?? report?.projectId ?? null;

  const reportDifficulty = normalizeMovement(report?.difficulty ?? "");

  const patchRows = (rows: any[]) =>
    (rows || []).map((row: any) => ({
      ...row,
      __report_difficulty: reportDifficulty,
      __report_category: report?.category ?? "",
      __report_description: report?.description ?? "",
      __report_remarks_action: report?.remarks_action ?? "",
    }));

  // 1) Highest priority: points saved directly for this report.
  // Newly added manual points are inserted into report_path_points, so check it first.
  try {
    const directTables = ["report_path_points", "report_points", "report_items"];
    const directFks = ["report_id", "reportId"];

    for (const table of directTables) {
      const probe = await supabase.from(table).select("*").limit(1);
      if (probe.error) continue;

      for (const fk of directFks) {
        let query = supabase.from(table).select("*").eq(fk, reportId);
        if (["report_path_points", "report_points", "report_items", "route_points"].includes(table)) {
          query = query.order("seq", { ascending: true, nullsFirst: false });
        }
        const { data, error } = await query;
        if (!error && Array.isArray(data) && data.length) {
          return { points: patchRows(data), report, routeId };
        }
      }
    }
  } catch {}

  // 2) Fallback: legacy/mixed sources. Keep report-linked lookups ahead of route/project lookups.
  for (const table of TABLE_CANDIDATES) {
    try {
      const probe = await supabase.from(table).select("*").limit(1);
      if (probe.error) continue;

      const orderedFks = [...FK_CANDIDATES].sort((a, b) => {
        const rank = (fk: string) => {
          const t = fk.toLowerCase();
          if (t.includes("report")) return 0;
          if (t.includes("route")) return 1;
          if (t.includes("project")) return 2;
          return 3;
        };
        return rank(a) - rank(b);
      });

      for (const fk of orderedFks) {
        const targetValue = fk.toLowerCase().includes("route")
          ? routeId
          : fk.toLowerCase().includes("project")
            ? projectId
            : reportId;

        if (!targetValue) continue;

        let query = supabase.from(table).select("*").eq(fk, targetValue);
        if (["report_path_points", "report_points", "report_items", "route_points"].includes(table)) {
          query = query.order("seq", { ascending: true, nullsFirst: false });
        }
        const { data, error } = await query;

        if (!error && Array.isArray(data) && data.length) {
          return { points: patchRows(data), report, routeId };
        }
      }
    } catch {}
  }

  // 3) Last fallback: single lat/lon on the report row itself.
  if (report?.loc_lat && report?.loc_lon) {
    return {
      points: [
        {
          loc_lat: report.loc_lat,
          loc_lon: report.loc_lon,
          details: report?.category ?? "",
          location: "",
          __report_difficulty: reportDifficulty,
          __report_category: report?.category ?? "",
          __report_description: report?.description ?? "",
          __report_remarks_action: report?.remarks_action ?? "",
        },
      ],
      report,
      routeId,
    };
  }

  throw new Error(
    `Points not found.
report_id=${reportId}
route_id=${routeId || "NULL"} project_id=${projectId || "NULL"}`
  );
}

/** =========================
 * BODY ROW BUILDER
 * ========================= */
async function makeBodyRow(supabase: any, p: NormalizedPoint, includePhotos: boolean) {
  let lat = p.__lat ?? null;
  let lon = p.__lon ?? null;

  if ((lat == null || lon == null) && p.ne_coordinate) {
    const parsed = parseNEToDecimal(p.ne_coordinate);
    if (parsed) {
      lat = parsed.lat;
      lon = parsed.lon;
    }
  }

  let locText = (p.location || "").trim();
  if (lat != null && lon != null && Number.isFinite(lat) && Number.isFinite(lon)) {
    const place = await reverseGeocodeOSM(lat, lon);
    if (place) {
      locText = place;
    } else if (!locText) {
      locText = fallbackLocationFromLatLon(lat, lon);
    }
  }
  if (!locText) locText = "—";

  return new TableRow({
    cantSplit: true,
children: [
      textCell(p.gps_no, 1, AlignmentType.CENTER, VerticalAlign.CENTER),
      textCell(p.kms, 1, AlignmentType.CENTER, VerticalAlign.CENTER),
      textCell(p.ne_coordinate, 1, AlignmentType.CENTER, VerticalAlign.CENTER),

      await detailsCellWithIcon(p.details, 2, VerticalAlign.CENTER),
      textCell(locText, 2, AlignmentType.LEFT, VerticalAlign.CENTER),
      textCell(p.description || "—", 1, AlignmentType.LEFT, VerticalAlign.CENTER),

      await movementCell(p.movement),

      await photoCell(supabase, p.photo_refs, includePhotos, p.photo_description),
    ],
  });
}


/** =========================
 * ✅ FULL-PAGE PHOTO LAYOUT (one photo per page)
 * Top info table (green header) + big centered photo
 * ========================= */
function getPhotoTheme(movement: string) {
  const mm = normalizeMovement(movement);
  if (mm === "green") return PHOTO_THEME.green;
  if (mm === "yellow") return PHOTO_THEME.yellow;
  if (mm === "red") return PHOTO_THEME.red;
  return PHOTO_THEME.default;
}

function photoPageHeaderCell(
  text: string,
  widthPct: number,
  fillColor: string,
  textColor: string = "FFFFFF"
) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: fillColor },
    borders: CELL_BORDERS,
    margins: { top: 220, bottom: 220, left: 160, right: 160 } as any,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: STYLE.spacing.none,
        children: [new TextRun({ text, bold: true, color: textColor, size: 32 })],
      }),
    ],
  });
}

function photoPageValueCell(
  text: string,
  widthPct: number,
  align: AlignmentType = AlignmentType.LEFT,
  fillColor: string = PHOTO_PAGE_ROW_FILL,
  textColor: string = PHOTO_THEME.default.text
) {
  const lines = splitLines(text);
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: fillColor },
    borders: CELL_BORDERS,
    margins: { top: 140, bottom: 140, left: 180, right: 180 } as any,
    children: lines.map((ln) =>
      new Paragraph({
        alignment: align,
        spacing: STYLE.spacing.cell,
        children: [new TextRun({ text: ln, bold: true, size: 32, color: textColor })],
      })
    ),
  });
}

/** ✅ NEW: Photo-page helpers (OBS icon + Route difficulty) */
function routeDifficultyFill(m: string) {
  const mm = normalizeMovement(m);
  if (mm === "green") return "C6EFCE"; // light green
  if (mm === "yellow") return "FFF2CC"; // light yellow
  if (mm === "red") return "F8CBAD"; // light red
  return PHOTO_PAGE_ROW_FILL;
}

function photoPageRouteDifficultyCell(
  movement: string,
  widthPct: number,
  fillColor?: string
) {
  const mm = normalizeMovement(movement);
  const label = mm ? mm.toUpperCase() : "—";

  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: fillColor || routeDifficultyFill(mm) },
    borders: CELL_BORDERS,
    margins: { top: 140, bottom: 140, left: 180, right: 180 } as any,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: STYLE.spacing.cell,
        children: [new TextRun({ text: label, bold: true, size: 32, color: "0B3D2E" })],
      }),
    ],
  });
}


async function photoPageCategoryCell(
  detailsText: string,
  widthPct: number,
  fillColor: string
) {
  const kind = detectDetailKind(detailsText || "");
  const iconBytes = kind ? await getDocxCategoryIcon(kind) : null;
  const categoryText = String(detailsText || "").trim() || "—";

  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

  const inner = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 28, type: WidthType.PERCENTAGE },
            borders: { top: none, bottom: none, left: none, right: none },
            verticalAlign: VerticalAlign.TOP,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 80, after: 0 },
                children: iconBytes
                  ? [new ImageRun({ data: iconBytes, transformation: { width: 58, height: 58 } })]
                  : [new TextRun({ text: "", size: 2 })],
              }),
            ],
          }),
          new TableCell({
            width: { size: 72, type: WidthType.PERCENTAGE },
            borders: { top: none, bottom: none, left: none, right: none },
            verticalAlign: VerticalAlign.TOP,
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: STYLE.spacing.cell,
                children: [new TextRun({ text: categoryText, bold: true, size: 32, color: "0B3D2E" })],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: fillColor },
    borders: CELL_BORDERS,
    margins: { top: 120, bottom: 120, left: 120, right: 120 } as any,
    children: [inner],
  });
}

function photoPageObservationCell(
  descText: string,
  widthPct: number,
  fillColor: string
) {
  const lines = splitLines(String(descText || "").trim() || "—");
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: fillColor },
    borders: CELL_BORDERS,
    margins: { top: 140, bottom: 140, left: 180, right: 180 } as any,
    children: lines.map((ln) =>
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: STYLE.spacing.cell,
        children: [new TextRun({ text: ln, bold: true, size: 32, color: "0B3D2E" })],
      })
    ),
  });
}

async function photoPageObsCell(detailsText: string, descText: string, widthPct: number) {
  const kind = detectDetailKind(detailsText || "");
  // High-res source so the icon stays sharp even when displayed
  const iconBytes = kind ? await getDocxCategoryIcon(kind) : null;

  const d1 = String(detailsText || "").trim();
  const d2 = String(descText || "").trim();

  // ✅ Only include description if it exists (prevents stray "—" line)
  const combined = [d1 || "—", d2].filter(Boolean).join("\n");
  const lines = splitLines(combined);

  const first = lines[0] ?? "—";
  const rest = lines.slice(1).filter((x) => String(x).trim() !== "—");

  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

  // ✅ Perfect alignment: use a 2-column inner table (icon | text)
  const iconTextRow = new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 12, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: { top: none, bottom: none, left: none, right: none },
            margins: { top: 0, bottom: 0, left: 0, right: 120 } as any,
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                spacing: STYLE.spacing.none,
                children: iconBytes
                  ? [new ImageRun({ data: iconBytes, transformation: { width: 46, height: 46 } })]
                  : [],
              }),
            ],
          }),
          new TableCell({
            width: { size: 88, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            borders: { top: none, bottom: none, left: none, right: none },
            margins: { top: 0, bottom: 0, left: 0, right: 0 } as any,
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: STYLE.spacing.none,
                children: [new TextRun({ text: first, bold: true, size: 32, color: "0B3D2E" })],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  const restParas = rest.map(
    (ln) =>
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: STYLE.spacing.cell,
        children: [new TextRun({ text: ln, bold: true, size: 32, color: "0B3D2E" })],
      })
  );

  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: PHOTO_PAGE_ROW_FILL },
    borders: CELL_BORDERS,
    margins: { top: 140, bottom: 140, left: 180, right: 180 } as any,
    children: [
      new Paragraph({ spacing: STYLE.spacing.none, text: "" }),
      iconTextRow as any,
      ...restParas,
    ],
  });
}





// =========================
// LAST PAGE: SUMMARY LIST TABLES (Bridges / Toll Plazas / Metro Sites / Tree Branches)
// Added WITHOUT changing existing data logic — uses already-normalized points.
// =========================
function movementRemark(movement: string) {
  const mm = normalizeMovement(movement);
  if (mm === "green") return "Clear Pass for TBM";
  if (mm === "yellow") return "Proceed with Caution";
  if (mm === "red") return "Not Recommended";
  return "—";
}

function summaryItemDescription(p: NormalizedPoint) {
  const title = String(p.details || "").trim() || "—";
  const desc = String(p.description || "").trim();
  const loc = String(p.location || "").trim();

  const parts: string[] = [];
  const first = loc ? `${title} - ${loc}` : title;
  parts.push(first);

  const cleanDesc = cleanExportDescription(title, desc);
  if (cleanDesc && cleanDesc.toLowerCase() !== loc.toLowerCase()) {
    parts.push(cleanDesc);
  }

  return parts.join("\n");
}

function makeSummaryTable(items: NormalizedPoint[]) {
  const headerFill = "F2F2F2";

  const hdrCell = (text: string) =>
    new TableCell({
      borders: CELL_BORDERS,
      shading: { type: ShadingType.CLEAR, fill: headerFill },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 220, bottom: 220, left: 180, right: 180 } as any,
      children: [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: STYLE.spacing.none,
          children: [new TextRun({ text, bold: true, size: 34, color: "000000" })],
        }),
      ],
    });

  const valueCell = (text: string, align: AlignmentType = AlignmentType.LEFT, boldFirstLine = false) => {
    const lines = splitLines(text || "—");
    return new TableCell({
      borders: CELL_BORDERS,
      verticalAlign: VerticalAlign.TOP,
      margins: { top: 220, bottom: 220, left: 180, right: 180 } as any,
      children: lines.map((t, i) =>
        new Paragraph({
          alignment: align,
          spacing: { before: i === 0 ? 0 : 90, after: 0 } as any,
          children: [
            new TextRun({
              text: String(t || "—"),
              size: 32,
              bold: boldFirstLine && i === 0,
              color: "000000",
            }),
          ],
        })
      ),
    });
  };

  const rows: TableRow[] = [
    new TableRow({
      height: { value: 1000, rule: HeightRule.ATLEAST },
      children: [hdrCell("Kms"), hdrCell("Description"), hdrCell("Remark/Recommendation")],
    }),
  ];

  for (const p of items) {
    rows.push(
      new TableRow({
        height: { value: 1200, rule: HeightRule.ATLEAST },
        children: [
          valueCell(String(p.kms || "—"), AlignmentType.CENTER),
          valueCell(summaryItemDescription(p), AlignmentType.LEFT, true),
          valueCell(movementRemark(p.movement), AlignmentType.LEFT),
        ],
      })
    );
  }

  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows,
    columnWidths: [2500, 15000, 6078],
  });
}

async function buildLastSummaryListsSection(params: {
  projectName: string;
  points: NormalizedPoint[];
  footerDate?: string | Date;
}) {
  const { projectName, points, footerDate } = params;

  // ✅ Display names for ALL kinds (add more anytime, but anything missing still shows)
  const KIND_LABEL: Record<string, string> = {
    footpath_bridge: "List of Footpath Bridges",
    bridge: "List of Bridges",
    underpass: "List of Underpass",
    lt_cable: "List of LT Cables",
    ht_cable: "List of HT Cables",
    towerline_cable: "List of Towerline Cables",
    towerline: "List of Towerlines",
    junction_left: "List of Left Junctions / Turns",
    junction_right: "List of Right Junctions / Turns",
    bend: "List of Bends",
    diversion: "List of Diversions",
    tree: "List of Tree Branches",
    petrol: "List of Petrol Bunks",
    electric_sign: "List of Electric Sign Boards",
    signboard: "List of Sign Boards",
    camera_pole: "List of Camera Poles",
    toll: "List of Toll Plazas",
  };

  // ✅ Ordering (you can rearrange; anything not listed will appear later)
  const ORDER: string[] = [
    "bridge",
    "footpath_bridge",
    "underpass",
    "toll",
    "metro", // keyword category
    "diversion",
    "petrol",
    "electric_sign",
    "signboard",
    "camera_pole",
    "junction_left",
    "junction_right",
    "bend",
    "lt_cable",
    "ht_cable",
    "towerline_cable",
    "towerline",
    "tree",
  ];

  // ✅ Collect items per category
  const buckets = new Map<string, NormalizedPoint[]>();
  const ensure = (k: string) => {
    if (!buckets.has(k)) buckets.set(k, []);
    return buckets.get(k)!;
  };

  for (const p of points) {
    const details = String(p.details || "");
    const desc = String(p.description || "");
    const loc = String(p.location || "");
    const tAll = `${details} ${desc} ${loc}`.toLowerCase();

    // Special keyword bucket for metro (since detectDetailKind doesn’t return "metro")
    if (tAll.includes("metro")) {
      ensure("metro").push(p);
      continue;
    }

    const k = detectDetailKind(details);

    if (k) {
      ensure(k).push(p);
    } else {
      ensure("other").push(p);
    }
  }

  const children: any[] = [];
  children.push(new Paragraph({ spacing: { before: 120, after: 0 }, text: "" }));

  // Helper to print a section
  const addSection = (key: string) => {
    const items = buckets.get(key);
    if (!items || !items.length) return;

    // Bridges table in your reference shows without heading first — keep that behavior if you want:
    const isFirstBridgeLike =
      key === "bridge" && items.length > 0 && children.length <= 2; // after initial spacer

    const titleBase =
      key === "metro"
        ? "List of Metro Site"
        : (KIND_LABEL[key] || `List of ${key.replace(/_/g, " ")}`);

    const title = `${titleBase} (${items.length})`;

    if (!isFirstBridgeLike) {
      children.push(
        new Paragraph({
          spacing: { before: 60, after: 120 } as any,
          // User-required: page titles = 24pt => 48 half-points
          children: [new TextRun({ text: title, bold: true, size: 48, color: "000000" })],
        })
      );
    } else {
      // If you still want “no heading for first bridges table”, just skip title.
      // If you DO want the title, comment this else block and keep the if above only.
    }

    children.push(makeSummaryTable(items));
    children.push(new Paragraph({ spacing: { before: 260, after: 0 }, text: "" }));
  };

  // ✅ Print sections in preferred order
  for (const k of ORDER) addSection(k);

  // ✅ Print any remaining categories not in ORDER
  const remaining = Array.from(buckets.keys()).filter((k) => !ORDER.includes(k));
  for (const k of remaining) addSection(k);

  // ✅ If nothing to show
  const totalItems = Array.from(buckets.values()).reduce((a, b) => a + b.length, 0);
  if (!totalItems) {
    children.push(
      new Paragraph({
        spacing: { before: 200, after: 0 } as any,
        children: [new TextRun({ text: "No summary items available.", size: 32, color: "000000" })],
      })
    );
  }

  return {
    properties: sectionPropsA3Landscape(),
    footers: { default: buildFooterTablePages(footerDate ?? new Date()) },
    children,
  };
}


async function buildCategoryCountSummarySection(params: {
  projectName: string;
  points: NormalizedPoint[];
  footerDate?: string | Date;
}) {
  const { points, footerDate } = params;

  // Same labels used in "List of ..." page (kept in-sync)
  const KIND_LABEL: Record<string, string> = {
    footpath_bridge: "List of Footpath Bridges",
    bridge: "List of Bridges",
    underpass: "List of Underpass",
    lt_cable: "List of LT Cables",
    ht_cable: "List of HT Cables",
    towerline_cable: "List of Towerline Cables",
    towerline: "List of Towerlines",
    junction_left: "List of Left Junctions / Turns",
    junction_right: "List of Right Junctions / Turns",
    bend: "List of Bends",
    diversion: "List of Diversions",
    tree: "List of Tree Branches",
    petrol: "List of Petrol Bunks",
    electric_sign: "List of Electric Sign Boards",
    signboard: "List of Sign Boards",
    camera_pole: "List of Camera Poles",
    toll: "List of Toll Plazas",
  };

  const ORDER: string[] = [
    "bridge",
    "footpath_bridge",
    "underpass",
    "toll",
    "metro",
    "diversion",
    "petrol",
    "electric_sign",
    "signboard",
    "camera_pole",
    "junction_left",
    "junction_right",
    "bend",
    "lt_cable",
    "ht_cable",
    "towerline_cable",
    "towerline",
    "tree",
  ];

  const counts = new Map<string, number>();
  const inc = (k: string) => counts.set(k, (counts.get(k) || 0) + 1);

  for (const p of points) {
    const details = String(p.details || "");
    const desc = String(p.description || "");
    const loc = String(p.location || "");
    const tAll = `${details} ${desc} ${loc}`.toLowerCase();

    if (tAll.includes("metro")) {
      inc("metro");
      continue;
    }

    const k = detectDetailKind(details);
    if (k) inc(k);
    else inc("other");
  }

  const labelOf = (key: string) => {
    if (key === "metro") return "Metro Site";
    const raw = KIND_LABEL[key] || `List of ${key.replace(/_/g, " ")}`;
    return raw.replace(/^List of\s+/i, "").trim();
  };

  const headerFill = "F2F2F2";

  const hdr = (text: string) =>
    new TableCell({
      borders: CELL_BORDERS,
      shading: { type: ShadingType.CLEAR, fill: headerFill },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 240, bottom: 240, left: 220, right: 220 } as any,
      children: [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: STYLE.spacing.none,
          children: [new TextRun({ text, bold: true, size: 34, color: "000000" })],
        }),
      ],
    });

  const cell = (text: string, align: AlignmentType = AlignmentType.LEFT, bold = false) =>
    new TableCell({
      borders: CELL_BORDERS,
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 240, bottom: 240, left: 220, right: 220 } as any,
      children: [
        new Paragraph({
          alignment: align,
          spacing: STYLE.spacing.none,
          children: [new TextRun({ text, size: 32, bold, color: "000000" })],
        }),
      ],
    });

  const rows: TableRow[] = [
    new TableRow({
      height: { value: 1000, rule: HeightRule.ATLEAST },
      children: [hdr("Category"), hdr("Count")],
    }),
  ];

  const addRow = (key: string) => {
    const c = counts.get(key) || 0;
    if (!c) return;
    rows.push(
      new TableRow({
        height: { value: 950, rule: HeightRule.ATLEAST },
        children: [cell(labelOf(key), AlignmentType.LEFT, true), cell(String(c), AlignmentType.CENTER, true)],
      })
    );
  };

  for (const k of ORDER) addRow(k);

  // Any remaining non-ordered categories (rare, but safe)
  const remaining = Array.from(counts.keys()).filter((k) => !ORDER.includes(k) && (counts.get(k) || 0) > 0);
  remaining.sort();
  for (const k of remaining) addRow(k);

  const children: any[] = [];

  children.push(
    new Paragraph({
      spacing: { before: 160, after: 220 } as any,
      children: [new TextRun({ text: "CATEGORY STAGE SUMMARY", bold: true, size: 48, color: "000000" })],
    })
  );

  children.push(
    new Table({
      layout: TableLayoutType.FIXED,
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows,
      columnWidths: [17500, 6078],
    })
  );

  return {
    properties: sectionPropsA3Landscape(),
    footers: { default: buildFooterTablePages(footerDate ?? new Date()) },
    children,
  };
}


async function buildPhotoPageSection(
  params: {
    supabase: any;
    projectName: string;
    p: NormalizedPoint;
    photoRefs: string[];
    footerDate?: string | Date;
    watermarkEnabled?: boolean;
  }
) {
  const { supabase, projectName, p, photoRefs, footerDate, watermarkEnabled } = params;

  // Location name (reverse-geocode if possible)
  let lat = p.__lat ?? null;
  let lon = p.__lon ?? null;

  if ((lat == null || lon == null) && p.ne_coordinate) {
    const parsed = parseNEToDecimal(p.ne_coordinate);
    if (parsed) {
      lat = parsed.lat;
      lon = parsed.lon;
    }
  }

  let locText = (p.location || "").trim();
  if (lat != null && lon != null && Number.isFinite(lat) && Number.isFinite(lon)) {
    const place = await reverseGeocodeOSM(lat, lon);
    if (place) {
      locText = place;
    } else if (!locText) {
      locText = fallbackLocationFromLatLon(lat, lon);
    }
  }
  if (!locText) locText = "—";

  const theme = getPhotoTheme(p.movement);
  const rowFill = theme.body;
  const headerFill = theme.header;
  const bodyTextColor = theme.text;

  const categoryCell = await photoPageCategoryCell(p.details || "—", 15, rowFill);
  const observationCell = photoPageObservationCell(p.description || "", 19, rowFill);
  const remarksCell = photoPageObservationCell(
    p.remarks_action || p.movement || "—",
    16,
    rowFill
  );

  const topInfo = new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          photoPageHeaderCell("GPS\nLOCATION", 18, headerFill, "000000"),
          photoPageHeaderCell("KM", 12, headerFill, "000000"),
          photoPageHeaderCell("LOCATION", 20, headerFill, "000000"),
          photoPageHeaderCell("CATEGORY", 15, headerFill, "000000"),
          photoPageHeaderCell("OBSERVATION", 19, headerFill, "000000"),
          photoPageHeaderCell("REMARKS / ACTION", 16, headerFill, "000000"),
        ],
      }),
      new TableRow({
        height: { value: 3200, rule: HeightRule.ATLEAST },
        children: [
          photoPageValueCell(p.ne_coordinate || "—", 18, AlignmentType.LEFT, rowFill, bodyTextColor),
          photoPageValueCell(p.kms || "—", 12, AlignmentType.CENTER, rowFill, bodyTextColor),
          photoPageValueCell(locText, 20, AlignmentType.LEFT, rowFill, bodyTextColor),
          categoryCell,
          observationCell,
          remarksCell,
        ],
      }),
    ],
  });

  const refs = (photoRefs || []).filter(Boolean).slice(0, 2);

  const bytesA = refs[0] ? await resolvePhotoBytesForDocx(supabase, refs[0], "page") : null;
  const bytesB = refs[1] ? await resolvePhotoBytesForDocx(supabase, refs[1], "page") : null;

const wrapperNone = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

  const content: any[] = [];
  content.push(new Paragraph({ spacing: { before: 0, after: 0 }, text: "" }));
content.push(topInfo);

  // small gap between table and image (keep them together)
  content.push(new Paragraph({ spacing: { before: 80, after: 0 }, text: "" }));

  if (bytesA && !bytesB) {
    // ✅ Single image (reduced height so it stays with the table on the same page)
    content.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [new ImageRun({ data: bytesA, transformation: fitTransform(bytesA, 1500, 560) })],
      })
    );
  } else if (bytesA && bytesB) {
    // ✅ Two images side-by-side
    const imgs = new Table({
      layout: TableLayoutType.FIXED,
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: { top: wrapperNone, bottom: wrapperNone, left: wrapperNone, right: wrapperNone },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 0 },
                  children: [new ImageRun({ data: bytesA, transformation: fitTransform(bytesA, 730, 420) })],
                }),
              ],
            }),
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: { top: wrapperNone, bottom: wrapperNone, left: wrapperNone, right: wrapperNone },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 0, after: 0 },
                  children: [new ImageRun({ data: bytesB, transformation: fitTransform(bytesB, 730, 420) })],
                }),
              ],
            }),
          ],
        }),
      ],
    });
    content.push(imgs);
  } else {
    content.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 160, after: 0 },
        children: [new TextRun({ text: "Photo not available.", size: 32, bold: true, color: "B42318" })],
      })
    );
  }

  // Wrap table+image in one row so Word won't split them across pages
  const children: any[] = [
    new Table({
      layout: TableLayoutType.FIXED,
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              borders: { top: wrapperNone, bottom: wrapperNone, left: wrapperNone, right: wrapperNone },
              margins: { top: 0, bottom: 0, left: 0, right: 0 } as any,
              children: content,
            }),
          ],
        }),
      ],
    }),
  ];

  return {
    properties: {
      verticalAlign: VerticalAlign.CENTER,
      page: {
        size: { width: A3_W, height: A3_H, orientation: PageOrientation.LANDSCAPE },
        margin: TABLE_MARGIN as any,
        ...(pageBordersTSPL() as any),
      } as any,
    },
    footers: { default: buildFooterTablePages(footerDate ?? new Date()) },
    children,
  };
}


/** =========================
 * DOC builder
 * ========================= */




function __decodeHtmlEntities(s: string) {
  return String(s || "")
    .replace(/&nbsp;|&#160;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&ndash;/gi, "–")
    .replace(/&mdash;/gi, "—")
    .replace(/\u00A0/g, " ");
}

function htmlToDocxParagraphs_Conclusion(html: string): Paragraph[] {
  const safe = __decodeHtmlEntities(String(html || "").trim());
  if (!safe) return [];

  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div>${safe}</div>`, "text/html");
  const root = doc.body.firstElementChild as HTMLElement | null;
  if (!root) return [];

  const paragraphs: Paragraph[] = [];

  const base = {
    bold: false,
    italics: false,
    underline: false,
    size: 28, // 14pt
    align: AlignmentType.LEFT as AlignmentType,
  };

  const merge = (a: any, b: any) => ({
    bold: a.bold || b.bold,
    italics: a.italics || b.italics,
    underline: a.underline || b.underline,
    size: b.size ?? a.size,
    align: b.align ?? a.align,
  });

  const cssSizeToHalfPoints = (v: string) => {
    const s = String(v || "").trim().toLowerCase();
    if (!s) return null;
    const m = s.match(/^([0-9.]+)\s*(pt|px)?$/);
    if (!m) return null;
    const n = Number(m[1]);
    if (!Number.isFinite(n) || n <= 0) return null;
    const unit = (m[2] || "px") as "pt" | "px";
    const pt = unit === "px" ? n * 0.75 : n;
    return Math.round(pt * 2);
  };

  const walkInline = (node: Node, style: any, out: (TextRun | ExternalHyperlink)[]) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const t = __decodeHtmlEntities(node.nodeValue || "");
      if (t) {
        out.push(
          new TextRun({
            text: t,
            bold: style.bold || undefined,
            italics: style.italics || undefined,
            underline: style.underline ? { type: UnderlineType.SINGLE } : undefined,
            size: style.size ?? 28,
            color: "111111",
          })
        );
      }
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;

    const el = node as HTMLElement;
    const tag = el.tagName.toLowerCase();

    if (tag === "br") {
      out.push(new TextRun({ text: "\n" }));
      return;
    }

    // hyperlink
    if (tag === "a") {
      const href = el.getAttribute("href") || "";
      const runs: TextRun[] = [];
      const innerStyle = merge(style, { underline: true });
      for (const ch of Array.from(el.childNodes)) {
        // only TextRuns inside link
        if (ch.nodeType === Node.TEXT_NODE) {
          const t = __decodeHtmlEntities(ch.nodeValue || "");
          if (t) {
            runs.push(
              new TextRun({
                text: t,
                bold: innerStyle.bold || undefined,
                italics: innerStyle.italics || undefined,
                underline: { type: UnderlineType.SINGLE },
                size: innerStyle.size ?? 28,
                color: "0563C1",
              })
            );
          }
        } else {
          // recurse but force TextRun output
          const tmp: any[] = [];
          walkInline(ch, innerStyle, tmp);
          for (const rr of tmp) {
            if (rr instanceof TextRun) runs.push(rr);
          }
        }
      }
      if (href && runs.length) {
        out.push(new ExternalHyperlink({ link: href, children: runs }));
      } else {
        // fallback to normal rendering
        for (const r of runs) out.push(r);
      }
      return;
    }

    const computed: any = {};
    if (tag === "strong" || tag === "b") computed.bold = true;
    if (tag === "em" || tag === "i") computed.italics = true;
    if (tag === "u") computed.underline = true;

    const hp = cssSizeToHalfPoints((el.style as any)?.fontSize || "");
    if (hp) computed.size = hp;

    const mergedStyle = merge(style, computed);

    for (const ch of Array.from(el.childNodes)) {
      walkInline(ch, mergedStyle, out);
    }
  };

  const pushParagraph = (runs: (TextRun | ExternalHyperlink)[], style: any, extra?: any) => {
    paragraphs.push(
      new Paragraph({
        alignment: style.align ?? AlignmentType.LEFT,
        spacing: { before: 0, after: 160, line: 360 } as any,
        children: runs.length ? (runs as any) : [new TextRun({ text: "", size: style.size ?? 28 })],
        ...extra,
      })
    );
  };

  const block = (el: HTMLElement, style: any, listPrefix?: string) => {
    const tag = el.tagName.toLowerCase();

    if (tag === "li") {
      const runs: (TextRun | ExternalHyperlink)[] = [];
      const prefix = (listPrefix || "•") + "  ";
      runs.push(new TextRun({ text: prefix, size: style.size ?? 28, color: "111111" }));
      for (const ch of Array.from(el.childNodes)) walkInline(ch, style, runs);
      pushParagraph(runs, style, { indent: { left: 720, hanging: 360 } as any });
      return;
    }

    if (tag === "ul") {
      for (const li of Array.from(el.children)) {
        if ((li as HTMLElement).tagName.toLowerCase() === "li") block(li as HTMLElement, style, "•");
      }
      return;
    }

    if (tag === "ol") {
      let i = 1;
      for (const li of Array.from(el.children)) {
        if ((li as HTMLElement).tagName.toLowerCase() === "li") {
          block(li as HTMLElement, style, `${i}.`);
          i += 1;
        }
      }
      return;
    }

    if (tag === "p" || tag === "div") {
      const runs: (TextRun | ExternalHyperlink)[] = [];
      const alignCss = String((el.style as any)?.textAlign || "").toLowerCase();
      const align =
        alignCss === "center"
          ? AlignmentType.CENTER
          : alignCss === "right"
            ? AlignmentType.RIGHT
            : AlignmentType.LEFT;

      const fs = cssSizeToHalfPoints(String((el.style as any)?.fontSize || ""));
      const st = merge(style, { align, size: fs ?? undefined });

      for (const ch of Array.from(el.childNodes)) walkInline(ch, st, runs);
      // keep empty line
      if (!runs.length) runs.push(new TextRun({ text: "" }));
      pushParagraph(runs, st);
      return;
    }

    // fallback recurse
    for (const ch of Array.from(el.children)) block(ch as HTMLElement, style);
  };

  // Walk top-level nodes. Convert block elements to paragraphs, and also capture stray inline/text at root.
  let pendingInline: (TextRun | ExternalHyperlink)[] = [];

  const flushInline = () => {
    if (pendingInline.length) {
      pushParagraph(pendingInline, base);
      pendingInline = [];
    }
  };

  const isBlockTag = (tag: string) =>
    ["p", "div", "ul", "ol", "li", "h1", "h2", "h3", "h4", "h5", "h6"].includes(tag);

  for (const n of Array.from(root.childNodes)) {
    if (n.nodeType === Node.TEXT_NODE) {
      const t = __decodeHtmlEntities(n.nodeValue || "");
      if (t.trim()) {
        // treat as inline text; keep spaces
        walkInline(n, base, pendingInline);
      }
      continue;
    }

    if (n.nodeType !== Node.ELEMENT_NODE) continue;
    const el = n as HTMLElement;
    const tag = el.tagName.toLowerCase();

    if (tag.startsWith("h") && tag.length === 2) {
      flushInline();
      const runs: (TextRun | ExternalHyperlink)[] = [];
      const st = merge(base, { bold: true, size: 56 }); // 28pt
      for (const ch of Array.from(el.childNodes)) walkInline(ch, st, runs);
      pushParagraph(runs, st, { spacing: { before: 0, after: 220, line: 360 } as any });
      continue;
    }

    if (isBlockTag(tag)) {
      flushInline();
      block(el, base);
      continue;
    }

    // Inline element at root (e.g., <strong>text</strong>)
    walkInline(el, base, pendingInline);
  }

  flushInline();

  return paragraphs;
}

function buildConclusionSection(opts: { conclusionHtml: string; footerDate?: string | Date }) {
  const htmlRaw = String(opts.conclusionHtml || "").trim();
  if (!htmlRaw) return null;

  // Convert HTML -> plain text with reliable line breaks, then format to match template page.
  const text = __decodeHtmlEntities(htmlRaw)
    // blocks
    .replace(/<\/p>/gi, "\n")
    .replace(/<p[^>]*>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/li>/gi, "\n")
    .replace(/<li[^>]*>/gi, "• ")
    .replace(/<\/(ul|ol)>/gi, "\n")
    .replace(/<h[1-6][^>]*>/gi, "")
    .replace(/<\/h[1-6]>/gi, "\n")
    // strip remaining tags
    .replace(/<[^>]+>/g, "")
    // remove odd checkbox / square glyphs that TinyMCE sometimes inserts
    .replace(/[\u25A1\u25FB\u25FD\u25FE\u2610\u2611\u2612]+/g, "")
    .replace(/\r/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  // Split into logical lines
  let lines = text.split("\n").map((s) => s.replace(/\s+/g, " ").trim());

  // Remove empty lines at ends
  while (lines.length && !lines[0]) lines.shift();
  while (lines.length && !lines[lines.length - 1]) lines.pop();

  // Remove duplicate title if present in content
  const titleText = "Conclusion & Certification";
  if (lines.length && lines[0].toLowerCase() === titleText.toLowerCase()) lines.shift();

  // Normalize bullets: keep only bullet lines for the two main points
  const isBullet = (s: string) => s.startsWith("• ") || s.startsWith("- ") || s.startsWith("•");
  // Convert "-" bullets to "• "
  lines = lines.map((s) => (s.startsWith("- ") ? `• ${s.slice(2).trim()}` : s));

  const children: Paragraph[] = [];

  // Title
  children.push(
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 220 } as any,
      children: [new TextRun({ text: "Conclusion & Certification", bold: true, size: 64, color: "111111" })],
    })
  );

  // Build paragraphs with template spacing
  for (const ln0 of lines) {
    const ln = String(ln0 || "").trim();
    if (!ln) {
      children.push(new Paragraph({ spacing: { before: 0, after: 180 } as any, text: "" }));
      continue;
    }

    // Bold label rows exactly like screenshot
    const boldLine =
      ln === "Issued By" ||
      ln === "For and on behalf of:" ||
      ln.startsWith("RACE LBI") ||
      ln.startsWith("Date:") ||
      ln.startsWith("Authorized Contact:") ||
ln.startsWith("www.");

    if (isBullet(ln)) {
      const bulletText = ln.replace(/^•\s*/g, "").trim();
      children.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { before: 80, after: 120, line: 360 } as any,
          indent: { left: 720, hanging: 360 } as any,
          children: [
            new TextRun({ text: "•  ", size: 48, color: "111111" }),
            new TextRun({ text: bulletText, size: 48, color: "111111" }),
          ],
        })
      );
      continue;
    }

    // Normal paragraph
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 140, line: 360 } as any,
        children: [new TextRun({ text: ln, size: 48, bold: boldLine, color: "111111" })],
      })
    );
  }

  return {
    properties: sectionPropsA3Landscape(),
    footers: { default: buildFooterTablePages(opts.footerDate ?? new Date()) },
    children,
  };
}


async function buildDoc(opts: {
  supabase: any;
  projectId?: string;
  includePhotos: boolean;
  fileName: string;
  points: any[];
  extraPhotoRefs?: string[];
  watermark?: WatermarkOptions;
  autoSave?: boolean;
  footerDate?: string | Date;
  projectName?: string;
  cover?: CoverOptions;
}): Promise<Blob> {
  const rows: TableRow[] = [makeHeaderRow()];

  let normalized = enrichPointsAlways(opts.points);
  if (opts.extraPhotoRefs?.length) normalized = applyExtraPhotos(normalized, opts.extraPhotoRefs);

  for (const p of normalized) rows.push(await makeBodyRow(opts.supabase, p, opts.includePhotos));

  const table = new Table({
    style: "Table Grid Light1",
    layout: TableLayoutType.FIXED,
    width: { size: TABLE_TOTAL_W, type: WidthType.DXA },
    columnWidths: GRID_COLS,
    rows,
    alignment: AlignmentType.CENTER as any,
    indent: { size: 0, type: WidthType.DXA } as any,
  });

  const wmEnabled = !!opts.watermark?.enabled;
  const projectName = (opts.projectName || "PROJECT").trim();

  const coverEnabled = opts.cover?.enabled ?? true;
  const cover: CoverOptions = {
    enabled: coverEnabled,
    // Default to /images/logo_v2.png (common Next.js public/images placement)
    logoUrl: opts.cover?.logoUrl ?? "/images/logo_v2.png",
    logoWidth: opts.cover?.logoWidth ?? 390,
    // Slightly shorter height to avoid stretching in Word header rendering
    logoHeight: opts.cover?.logoHeight ?? 68,
    rightTopText: opts.cover?.rightTopText ?? "",
    topCenterText: opts.cover?.topCenterText ?? `${projectName} SUMMARY REPORT`,
    // Label inside the blue outlined box on the cover page
    recommendationText: opts.cover?.recommendationText ?? "SURVEY REPORT",
    footerLeftText: opts.cover?.footerLeftText ?? "Report by RACE Innovations Pvt ltd",
    footerEmail: opts.cover?.footerEmail ?? "kh@raceinnovations.in",
    footerWebsite: opts.cover?.footerWebsite ?? "https://raceinnovations.in/",
    datedLabel: opts.cover?.datedLabel ?? "Dated",
    date: opts.cover?.date ?? opts.footerDate ?? new Date(),
  };

  const sections: any[] = [];

  // Default logo-only header for all non-cover pages
  const headerDefault = await buildLogoOnlyHeader(cover.logoUrl, 220, 38);

  const headerPhotoTitle = await buildGATitleHeader({
    logoUrl: cover.logoUrl,
    projectName,
    includeGATitle: false,
    logoW: 220,
    logoH: 38,
  });


  // ✅ COVER PAGE (A3 Landscape)
  if (cover.enabled) {
    sections.push({
      properties: {
        verticalAlign: VerticalAlign.CENTER,
        page: {
          // A3 LANDSCAPE: do NOT swap width/height; swapping can make Word open as portrait.
          size: { width: A3_W, height: A3_H, orientation: PageOrientation.LANDSCAPE },
          margin: COVER_MARGIN as any,
          ...(pageBordersTSPL() as any),
        } as any,
      },
      headers: { default: await buildLogoOnlyHeader(cover.logoUrl, cover.logoWidth ?? 220, cover.logoHeight ?? 38) },
      footers: { default: buildCoverFooter(cover) },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 200 },
          children: [
            // User-required: First line title = 46pt => 92 half-points
            new TextRun({ text: String(cover.topCenterText || "").trim().toUpperCase(), size: 92, color: "1F4E79" }),
          ],
        }),
        // Keep the title block compact (reference uses less top whitespace).
        new Paragraph({ spacing: { before: 120, after: 0 }, text: "" }),
        coverTitleProject(projectName),
        new Paragraph({ spacing: { before: 140, after: 0 }, text: "" }),
        coverRecommendationBox(cover.recommendationText || "SURVEY REPORT"),
      ],
    });
  }

  // ✅ Objective+Map + GA pages (if available)
  try {
    const pid = String(opts.projectId || "").trim();
    if (pid) {
      const setup = await getProjectRouteSetup(opts.supabase, pid);
      if (setup) {
        const mapBytes = setup.routeMapUrl
          ? await resolvePhotoBytes(opts.supabase, setup.routeMapUrl)
          : null;
        const gaUrls = (setup.gaImageUrls || []).filter(Boolean);
        const gaBytesList: Uint8Array[] = [];
        for (const u of gaUrls) {
          try {
            const b = await resolvePhotoBytes(opts.supabase, u);
            if (b) gaBytesList.push(b);
          } catch {
            // ignore per-image failure
          }
        }

        const objectiveSec = await buildObjectiveRouteMapSection({
            projectName,
            objective: setup.objective || "—",
            routeMapBytes: mapBytes,
            routeLocations: (setup as any).locations || [],
            footerDate: opts.footerDate ?? new Date(),
          });
        objectiveSec.headers = { default: headerDefault };
        sections.push(objectiveSec);

        const gaSecs = await buildGADrawingSections({
          projectName,
          gaDrawingBytesList: gaBytesList,
          footerDate: opts.footerDate ?? new Date(),
        });
        // ✅ Keep ONLY titles at the top (near logo) without shifting the drawing.
        // Titles are moved into the header for GA pages.
        const gaHeaderFirst = await buildGATitleHeader({
          logoUrl: cover.logoUrl,
          projectName,
          includeGATitle: true,
          logoW: 220,
          logoH: 38,
        });
        const gaHeaderNext = await buildGATitleHeader({
          logoUrl: cover.logoUrl,
          projectName,
          includeGATitle: false,
          logoW: 220,
          logoH: 38,
        });

        for (let i = 0; i < gaSecs.length; i++) {
          const s = gaSecs[i] as any;
          s.headers = { default: i === 0 ? gaHeaderFirst : gaHeaderNext };
          sections.push(s);
        }



        // ✅ CATEGORY STAGE SUMMARY (after GA Drawing)
        try {
          const countSec = await buildCategoryCountSummarySection({
            projectName,
            points: normalized,
            footerDate: opts.footerDate ?? new Date(),
          });
          (countSec as any).headers = { default: headerDefault };
          sections.push(countSec);
        } catch {
          // ignore count-summary failures
        }

        // ✅ LIST OF FOOTPATH BRIDGES / BRIDGES / etc. (after Category Stage Summary)
        try {
          const lastSec = await buildLastSummaryListsSection({
            projectName,
            points: normalized,
            footerDate: opts.footerDate ?? new Date(),
          });
          (lastSec as any).headers = { default: headerDefault };
          sections.push(lastSec);
        } catch {
          // ignore last-page failures
        }



// ✅ FULL-PAGE PHOTO PAGES (1 image = full width, 2 images = side-by-side)
  try {
    for (const p of normalized) {
      const refs = (p.photo_refs || []).filter(Boolean).slice(0, 3);

      if (refs.length === 0) continue;
      if (refs.length === 1) {
        const photoSec = await buildPhotoPageSection({
          supabase: opts.supabase,
          projectName,
          p,
          photoRefs: [refs[0]],
          footerDate: opts.footerDate ?? new Date(),
          watermarkEnabled: wmEnabled,
        });
        photoSec.headers = { default: headerPhotoTitle };
        sections.push(photoSec);
      } else if (refs.length === 2) {
        const photoSec = await buildPhotoPageSection({
          supabase: opts.supabase,
          projectName,
          p,
          photoRefs: [refs[0], refs[1]],
          footerDate: opts.footerDate ?? new Date(),
          watermarkEnabled: wmEnabled,
        });
        photoSec.headers = { default: headerPhotoTitle };
        sections.push(photoSec);
      } else {
        // 3 photos: first 2 on one page, last one on next page
        const photoSec1 = await buildPhotoPageSection({
          supabase: opts.supabase,
          projectName,
          p,
          photoRefs: [refs[0], refs[1]],
          footerDate: opts.footerDate ?? new Date(),
          watermarkEnabled: wmEnabled,
        });
        photoSec1.headers = { default: headerPhotoTitle };
        sections.push(photoSec1);

        const photoSec2 = await buildPhotoPageSection({
          supabase: opts.supabase,
          projectName,
          p,
          photoRefs: [refs[2]],
          footerDate: opts.footerDate ?? new Date(),
          watermarkEnabled: wmEnabled,
        });
        photoSec2.headers = { default: headerPhotoTitle };
        sections.push(photoSec2);
      }
    }
  } catch {
    // ignore photo-page failures
  }


      }
    }
  } catch {
    // Do not break export if route setup fetch fails
  }

  // ✅ MAIN REPORT TABLE SECTION
  // The export modal count is based on reports selected/listed, so the DOCX must include
  // the actual report table section. Without this section, the file shows only cover / map /
  // GA / summary / photo pages.
  sections.push({
    properties: sectionPropsA3Landscape(),
    headers: { default: headerDefault },
    footers: { default: buildFooterTablePages(opts.footerDate ?? new Date()) },
    children: [table],
  });


  const projectId = String(opts.projectId || "").trim();

  // ✅ CONCLUSION & CERTIFICATION (last page)
  // Pulls HTML from project_route_pages.conclusion_html (latest non-null row).
  try {
    const { data: lastPage } = await opts.supabase
      .from("project_route_pages")
      .select("conclusion_html, created_at")
      .eq("project_id", projectId)
      .not("conclusion_html", "is", null)
      .order("created_at", { ascending: false })
      .limit(1)
      .maybeSingle();

    const conclusionHtml = String((lastPage as any)?.conclusion_html || "").trim();
    if (conclusionHtml) {
      const conclusionSec = buildConclusionSection({
        conclusionHtml,
        footerDate: cover.date ?? opts.footerDate ?? new Date(),
      });
      if (conclusionSec) sections.push(conclusionSec);
    }
  } catch {
    // ignore conclusion failures (do not block export)
  }


  const doc = new Document({ sections });

  // ✅ Build docx bytes then patch true Word page borders (client-style)
  const bytes = await Packer.toBuffer(doc);
  const patched = await applyRedPageBordersToDocxBytes(bytes as any);
  const blob = new Blob([patched], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });

  return blob;
}

/** =========================
 * EXPORTED DOCX
 * ========================= */
export async function downloadReportDOCX(supabase: any, reportId: string, opts: DownloadOpts = {}) {
  const includePhotos = opts.includePhotos ?? true;

  const { points, report } = await getPointsForReport(supabase, reportId);
  const extraPhotoRefs = includePhotos ? await getExtraPhotosForReport(supabase, reportId) : [];

  let projectName = "Project";
  let projectId: string | null = null;
  try {
    const pid = report?.project_id ?? report?.projectId;
    if (pid) {
      projectId = String(pid);
      const { data: project } = await supabase.from("projects").select("*").eq("id", pid).single();
      projectName = projectNameOf(project as any);
    }
  } catch {}

  await buildDoc({
    supabase,
    projectId: projectId || undefined,
    includePhotos,
    fileName: opts.fileName || `report-${String(reportId).slice(0, 8)}.docx`,
    points,
    extraPhotoRefs,
    watermark: opts.watermark,
    footerDate: report?.created_at || new Date(),
    projectName,
    cover: opts.cover,
  });
}


/** =========================
 * ✅ EXPORTED PROJECTS CSV
 * ========================= */
export async function downloadProjectsCSV(
  supabase: any,
  opts: { fileName?: string; fields?: string[] } = {}
) {
  const fileName = opts.fileName || "projects.csv";
  const fields = (opts.fields && opts.fields.length ? opts.fields : ["id", "name", "title", "project_name", "created_at"])
    .map(String)
    .filter(Boolean);

  // Pull projects
  const { data, error } = await supabase.from("projects").select("*").order("created_at", { ascending: false });
  if (error) throw error;

  const rows = Array.isArray(data) ? data : [];

  function csvEscape(v: any) {
    const s = (v === null || v === undefined) ? "" : String(v);
    const needs = /[",\n\r]/.test(s);
    const esc = s.replace(/"/g, '""');
    return needs ? `"${esc}"` : esc;
  }

  const header = fields.join(",");
  const body = rows
    .map((r: any) => fields.map((f) => csvEscape(r?.[f])).join(","))
    .join("\n");

  const csv = header + "\n" + body;
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  return { blob, fileName };
}



async function ensureRouteIdForMissingReportsOnServer(
  supabase: any,
  projectId: string,
  reportIds: string[]
) {
  if (!projectId || !reportIds.length) return;

  const { data: reports, error: rErr } = await supabase
    .from("reports")
    .select("id, route_id")
    .in("id", reportIds);
  if (rErr) throw rErr;

  const missingIds = (reports || [])
    .filter((r: any) => !r?.route_id)
    .map((r: any) => String(r.id))
    .filter(Boolean);

  if (!missingIds.length) return;

  const { data: latestRoute, error: routeErr } = await supabase
    .from("routes")
    .select("id")
    .eq("project_id", projectId)
    .order("created_at", { ascending: false })
    .limit(1)
    .maybeSingle();
  if (routeErr) throw routeErr;

  const routeId = latestRoute?.id;
  if (!routeId) return;

  const { error: updErr } = await supabase
    .from("reports")
    .update({ route_id: routeId })
    .in("id", missingIds);
  if (updErr) throw updErr;
}

export async function generateProjectDOCX(
  supabase: any,
  projectId: string,
  opts: DownloadOpts = {}
): Promise<{ blob: Blob; fileName: string }> {
  const includePhotos = opts.includePhotos ?? true;

  const { data: project, error: pErr } = await supabase
    .from("projects")
    .select("*")
    .eq("id", projectId)
    .single();
  if (pErr) throw pErr;

  const { data: reports, error: rErr } = await supabase
    .from("reports")
    .select("*")
    .eq("project_id", projectId)
    .order("created_at", { ascending: true });
  if (rErr) throw rErr;

  const allReportIds = ((reports || []) as ReportRow[]).map((r) => String(r.id)).filter(Boolean);
  await ensureRouteIdForMissingReportsOnServer(supabase, projectId, allReportIds);

  const collected: any[] = [];

  for (const r of (reports || []) as ReportRow[]) {
    let loaded: any = null;
    try {
      loaded = await getPointsForReport(supabase, r.id);
    } catch {
      continue;
    }

    const { points } = loaded || {};
    if (!Array.isArray(points) || !points.length) continue;

    const reportPoints = points.map((pt: any, idx: number) => ({
      ...pt,
      __report_difficulty: r?.difficulty ?? "",
      __report_category: r?.category ?? "",
      __report_description: r?.description ?? "",
      __report_remarks_action: r?.remarks_action ?? "",
      __point_order: pointSortValue(pt, idx),
      __report_created_at: r?.created_at ?? "",
      __report_id: r?.id ?? "",
    }));

    if (includePhotos) {
      const extra = await getExtraPhotosForReport(supabase, r.id);
      if (extra.length && reportPoints.length) {
        reportPoints[0].photo_refs = Array.from(
          new Set([...(reportPoints[0].photo_refs || []), ...extra])
        ).slice(0, 3);
      }
    }

    collected.push(...reportPoints);
  }

  collected.sort((a, b) => {
    const ao = Number(a.__point_order ?? 0);
    const bo = Number(b.__point_order ?? 0);
    if (Number.isFinite(ao) && Number.isFinite(bo) && ao !== bo) return ao - bo;
    return String(a.__report_created_at || "").localeCompare(String(b.__report_created_at || ""));
  });

  const allPoints = enrichPointsAlways(collected);

  const name = projectNameOf(project as ProjectRow);
  const fileName = opts.fileName || `${name}-ALL-REPORTS.docx`;

  const blob = await buildDoc({
    supabase,
    projectId,
    includePhotos,
    fileName,
    points: allPoints,
    watermark: opts.watermark,
    autoSave: false,
    footerDate: new Date(),
    projectName: name,
    cover: opts.cover,
  });

  return { blob, fileName };
}

export async function generateProjectDOCXByReportIds(
  supabase: any,
  projectId: string,
  reportIds: string[],
  opts: DownloadOpts = {}
): Promise<{ blob: Blob; fileName: string }> {
  const includePhotos = opts.includePhotos ?? true;

  const { data: project, error: pErr } = await supabase
    .from("projects")
    .select("*")
    .eq("id", projectId)
    .single();
  if (pErr) throw pErr;

  await ensureRouteIdForMissingReportsOnServer(supabase, projectId, reportIds);

  const collected: any[] = [];

  for (const reportId of reportIds) {
    let loaded: any = null;
    try {
      loaded = await getPointsForReport(supabase, reportId);
    } catch {
      continue;
    }

    const { points, report } = loaded || {};
    if (!Array.isArray(points) || !points.length) continue;

    const reportPoints = points.map((pt: any, idx: number) => ({
      ...pt,
      __report_difficulty: report?.difficulty ?? "",
      __report_category: report?.category ?? "",
      __report_description: report?.description ?? "",
      __report_remarks_action: report?.remarks_action ?? "",
      __point_order: pointSortValue(pt, idx),
      __report_created_at: report?.created_at ?? "",
      __report_id: report?.id ?? "",
    }));

    if (includePhotos) {
      const extra = await getExtraPhotosForReport(supabase, reportId);
      if (extra.length && reportPoints.length) {
        reportPoints[0].photo_refs = Array.from(
          new Set([...(reportPoints[0].photo_refs || []), ...extra])
        ).slice(0, 3);
      }
    }

    collected.push(...reportPoints);
  }

  collected.sort((a, b) => {
    const ao = Number(a.__point_order ?? 0);
    const bo = Number(b.__point_order ?? 0);
    if (Number.isFinite(ao) && Number.isFinite(bo) && ao !== bo) return ao - bo;
    return String(a.__report_created_at || "").localeCompare(String(b.__report_created_at || ""));
  });

  const allPoints = enrichPointsAlways(collected);

  const name = projectNameOf(project as ProjectRow);
  const fileName = opts.fileName || `${name}-REPORTS-${reportIds.length}.docx`;

  const blob = await buildDoc({
    supabase,
    projectId,
    includePhotos,
    fileName,
    points: allPoints,
    watermark: opts.watermark,
    autoSave: false,
    footerDate: new Date(),
    projectName: name,
    cover: opts.cover,
  });

  return { blob, fileName };
}

/** =========================
 * EXPORTED GPX
 * ========================= */
async function collectGpxPointsForReportId(supabase: any, reportId: string): Promise<GPXPoint[]> {
  let loaded: any = null;
  try {
    loaded = await getPointsForReport(supabase, reportId);
  } catch {
    return [];
  }

  const { points, report } = loaded || {};
  const norm = enrichPointsAlways(points || []);

  const out: GPXPoint[] = [];
  const baseTime = report?.created_at ? new Date(report.created_at) : new Date();

  let tick = 0;
  for (const p of norm) {
    let lat = p.__lat ?? null;
    let lon = p.__lon ?? null;

    if ((lat == null || lon == null) && p.ne_coordinate) {
      const parsed = parseNEToDecimal(p.ne_coordinate);
      if (parsed) {
        lat = parsed.lat;
        lon = parsed.lon;
      }
    }

    if (lat == null || lon == null) continue;
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) continue;

    const t = new Date(baseTime.getTime() + tick * 1000);
    tick += 2;

    out.push({ lat, lon, time: isoUtc(t) });
  }

  return out;
}

export async function generateProjectGPXByReportIds(
  supabase: any,
  projectId: string,
  reportIds: string[],
  opts: { fileName?: string; name?: string } = {}
): Promise<{ blob: Blob; fileName: string }> {
  const { data: project } = await supabase.from("projects").select("*").eq("id", projectId).single();
  const baseName = opts.name || projectNameOf(project as any);

  await ensureRouteIdForMissingReportsOnServer(supabase, projectId, reportIds);

  const points: GPXPoint[] = [];
  for (const rid of reportIds) {
    const pts = await collectGpxPointsForReportId(supabase, rid);
    points.push(...pts);
  }

  if (!points.length) throw new Error("No valid NE coordinate points found to export GPX.");

  const xml = toGpxXml({
    name: baseName,
    creator: "Recorded in TSPL Web App",
    points,
  });

  const fileName = opts.fileName || `${String(baseName).slice(0, 80)}.gpx`;
  const blob = new Blob([xml], { type: "application/gpx+xml" });

  return { blob, fileName };
}

export async function generateProjectGPX(
  supabase: any,
  projectId: string,
  opts: { fileName?: string; name?: string } = {}
): Promise<{ blob: Blob; fileName: string }> {
  const { data: project, error: pErr } = await supabase.from("projects").select("*").eq("id", projectId).single();
  if (pErr) throw pErr;

  const { data: reports, error: rErr } = await supabase
    .from("reports")
    .select("id")
    .eq("project_id", projectId)
    .order("created_at", { ascending: true });

  if (rErr) throw rErr;

  const ids = (reports || []).map((r: any) => r.id).filter(Boolean);
  if (!ids.length) throw new Error("No reports available for GPX export.");

  const name = opts.name || projectNameOf(project as any);
  return generateProjectGPXByReportIds(supabase, projectId, ids, {
    name,
    fileName: opts.fileName || `${String(name).slice(0, 80)}-ALL.gpx`,
  });
}


export async function generateProjectDOCXBuffer(
  supabase: any,
  projectId: string,
  opts: DownloadOpts = {}
): Promise<{ buffer: Buffer; fileName: string }> {
  const { blob, fileName } = await generateProjectDOCX(supabase, projectId, opts);
  const buffer = Buffer.from(await blob.arrayBuffer());
  return { buffer, fileName };
}

export async function generateProjectDOCXByReportIdsBuffer(
  supabase: any,
  projectId: string,
  reportIds: string[],
  opts: DownloadOpts = {}
): Promise<{ buffer: Buffer; fileName: string }> {
  const { blob, fileName } = await generateProjectDOCXByReportIds(supabase, projectId, reportIds, opts);
  const buffer = Buffer.from(await blob.arrayBuffer());
  return { buffer, fileName };
}
