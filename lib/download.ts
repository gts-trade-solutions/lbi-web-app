"use client";

/**
 * lib/download.ts (FULL FILE)
 * ✅ Fixes included:
 * 1) Uses `difficulty` column (NOT vehicle_movement)
 * 2) DOCX "LOCATION" column shows readable place name (Thomas Mount, Guindy, etc.)
 *    via reverse-geocoding from NE coordinates (OpenStreetMap Nominatim)
 * 3) Prevents "description showing in location" by keeping location separate
 * 4) Keeps vehicle movement square color based on difficulty (green/yellow/red)
 */

import { saveAs } from "file-saver";
import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeightRule,
  ImageRun,
  Packer,
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

/** =========================
 * GPX Types
 * ========================= */
type GPXPoint = { lat: number; lon: number; time?: string };

function isoUtc(d: Date) {
  return d.toISOString();
}

/**
 * Parses strings like:
 * "N28 02.912 E84 48.869"
 * "N28.12345 E84.98765"
 * "N28 02.912\nE84 48.869"
 */
function parseNEToDecimal(ne: string): { lat: number; lon: number } | null {
  const t = String(ne || "")
    .replace(/\r/g, " ")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!t) return null;

  const m = t.match(
    /([NS])\s*(\d{1,3}(?:\.\d+)?)\s*(?:(\d{1,3}(?:\.\d+)?))?\s*([EW])\s*(\d{1,3}(?:\.\d+)?)\s*(?:(\d{1,3}(?:\.\d+)?))?/i
  );
  if (!m) return null;

  const ns = m[1].toUpperCase();
  const latDeg = Number(m[2]);
  const latMin = m[3] != null ? Number(m[3]) : null;

  const ew = m[4].toUpperCase();
  const lonDeg = Number(m[5]);
  const lonMin = m[6] != null ? Number(m[6]) : null;

  if (!Number.isFinite(latDeg) || !Number.isFinite(lonDeg)) return null;

  let lat = latDeg;
  let lon = lonDeg;

  if (latMin != null && Number.isFinite(latMin)) lat = latDeg + latMin / 60;
  if (lonMin != null && Number.isFinite(lonMin)) lon = lonDeg + lonMin / 60;

  if (ns === "S") lat = -Math.abs(lat);
  if (ew === "W") lon = -Math.abs(lon);

  if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
  return { lat, lon };
}

/** =========================
 * ✅ Reverse Geocode (NE -> readable location)
 * Uses OpenStreetMap Nominatim reverse API
 * ========================= */
const REVERSE_CACHE = new Map<string, string>();
const REVERSE_INFLIGHT = new Map<string, Promise<string>>();
const REVERSE_TIMEOUT_MS = 9000;

function coordKey(lat: number, lon: number) {
  // round to ~100m (reduces repeated calls)
  const r = (n: number) => Math.round(n * 1000) / 1000;
  return `${r(lat)},${r(lon)}`;
}

function pickFirst(obj: any, keys: string[]) {
  for (const k of keys) {
    const v = obj?.[k];
    if (typeof v === "string" && v.trim()) return v.trim();
  }
  return "";
}

function formatOsmAddress(addr: any) {
  const p1 = pickFirst(addr, ["neighbourhood", "suburb", "quarter", "hamlet"]);
  const p2 = pickFirst(addr, ["city_district", "district", "borough", "county", "state_district"]);
  const p3 = pickFirst(addr, ["city", "town", "village", "municipality"]);

  const parts = [p1, p2, p3].filter(Boolean);

  if (parts.length < 2) {
    const st = pickFirst(addr, ["state"]);
    if (st) parts.push(st);
  }

  return parts.join(", ");
}

async function fetchWithTimeout(url: string, ms: number) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), ms);
  try {
    return await fetch(url, {
      signal: controller.signal,
      headers: { Accept: "application/json" },
    });
  } finally {
    clearTimeout(t);
  }
}

async function reverseGeocodeOSM(lat: number, lon: number): Promise<string> {
  const key = coordKey(lat, lon);
  if (REVERSE_CACHE.has(key)) return REVERSE_CACHE.get(key)!;
  if (REVERSE_INFLIGHT.has(key)) return await REVERSE_INFLIGHT.get(key)!;

  const job = (async () => {
    try {
      const url =
        `https://nominatim.openstreetmap.org/reverse` +
        `?format=jsonv2&lat=${encodeURIComponent(lat)}&lon=${encodeURIComponent(lon)}` +
        `&zoom=16&addressdetails=1`;

      const res = await fetchWithTimeout(url, REVERSE_TIMEOUT_MS);
      if (!res.ok) return "";

      const json: any = await res.json();
      const addr = json?.address || {};
      const label = formatOsmAddress(addr);

      const out =
        (label || "").trim() ||
        (json?.display_name ? String(json.display_name).split(",").slice(0, 3).join(",").trim() : "");

      if (out) REVERSE_CACHE.set(key, out);
      return out || "";
    } catch {
      return "";
    } finally {
      REVERSE_INFLIGHT.delete(key);
    }
  })();

  REVERSE_INFLIGHT.set(key, job);
  const result = await job;
  if (result) REVERSE_CACHE.set(key, result);
  return result;
}

/** =========================
 * GPX Generator
 * ========================= */
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
const TABLE_TOTAL_W = 15848;
// 11 physical columns (sum = 15848)
const GRID_COLS = [846, 1134, 1701, 2884, 1085, 906, 1787, 2607, 1111, 676, 1111];

const HEADER_FILL = "365F91";
const PAGE_BORDER_COLOR = "C00000";

const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const CELL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const P0 = { before: 0, after: 0 };

const DEBUG_PHOTOS = false;

/** ✅ Watermark opts */
export type WatermarkOptions = {
  enabled?: boolean;
  text?: string;
};

export type DownloadOpts = {
  includePhotos?: boolean;
  fileName?: string;
  watermark?: WatermarkOptions;
};

/** =========================
 * DB Types
 * ========================= */
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
  difficulty?: string | null; // ✅ correct column
};

/** =========================
 * Point Normalization
 * ========================= */
type NormalizedPoint = {
  gps_no: string;
  kms: string;
  ne_coordinate: string;
  details: string;
  location: string; // will be replaced by reverse-geocode if coords exist
  photo_refs: string[];
  movement: VehicleMovement;
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

/** Red TSPL page border */
function pageBordersTSPL(): any {
  return {
    borders: {
      pageBorders: {
        top: { style: BorderStyle.DOUBLE, size: 4, color: PAGE_BORDER_COLOR, space: 24 },
        left: { style: BorderStyle.DOUBLE, size: 4, color: PAGE_BORDER_COLOR, space: 24 },
        bottom: { style: BorderStyle.DOUBLE, size: 4, color: PAGE_BORDER_COLOR, space: 24 },
        right: { style: BorderStyle.DOUBLE, size: 4, color: PAGE_BORDER_COLOR, space: 24 },
      },
    },
  };
}

/** =========================
 * Text helpers
 * ========================= */
function run(text: string, opts?: { bold?: boolean; color?: string; size?: number; underline?: boolean }) {
  return new TextRun({
    text,
    bold: opts?.bold,
    color: opts?.color,
    underline: opts?.underline ? { type: UnderlineType.SINGLE } : undefined,
    size: opts?.size ?? 24,
  });
}

function paragraphPlain(text: string, align: AlignmentType) {
  return new Paragraph({
    alignment: align,
    spacing: P0,
    children: [run(text)],
  });
}

function paragraphFromLine(line: string) {
  const t = (line ?? "").toString().trimEnd();
  const isBullet = t.trim().startsWith("•") || t.trim().startsWith("-") || t.trim().startsWith("• ");

  if (!isBullet) return paragraphPlain(t, AlignmentType.LEFT);

  const normalized = t.trim().startsWith("-") ? `• ${t.trim().slice(1).trim()}` : t.trim();

  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: P0,
    indent: { left: 360, hanging: 180 },
    children: [run(normalized)],
  });
}

function splitLines(text: string) {
  const lines = (text || "").toString().split("\n").map((x) => x.trimEnd());
  const filtered = lines.filter((x) => x.length > 0);
  return filtered.length ? filtered : [""];
}

/** =========================
 * ✅ WATERMARK
 * ========================= */
function splitWatermarkText(full: string): { left: string; right: string } {
  const t = String(full || "").trim();
  const idx = t.indexOf(":");
  if (idx !== -1) {
    const left = t.slice(0, idx + 1);
    const right = t.slice(idx + 1).trim();
    return { left: left.trim(), right: right };
  }
  return { left: "CONFIDENTIAL REPORT:", right: t };
}

async function watermarkPngBytesDiagonal(text: string) {
  const W = 1600;
  const H = 900;
  const canvas = document.createElement("canvas");
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext("2d")!;
  ctx.clearRect(0, 0, W, H);

  ctx.save();
  ctx.translate(W / 2, H / 2);
  ctx.rotate((-30 * Math.PI) / 180);

  ctx.font = "700 160px Arial";
  ctx.fillStyle = "rgba(120,120,120,0.12)";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(text || "CONFIDENTIAL", 0, 0);
  ctx.restore();

  const base64 = canvas.toDataURL("image/png").split(",")[1];
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function buildHeaderWithDiagonalWatermark(wmText: string) {
  const bytes = await watermarkPngBytesDiagonal("CONFIDENTIAL");

  const watermarkImage = new ImageRun({
    data: bytes,
    transformation: { width: 900, height: 520 },
    floating: {
      horizontalPosition: {
        relative: HorizontalPositionRelativeFrom.PAGE,
        align: HorizontalPositionAlign.CENTER,
      },
      verticalPosition: {
        relative: VerticalPositionRelativeFrom.PAGE,
        align: VerticalPositionAlign.CENTER,
      },
      wrap: { type: TextWrappingType.NONE },
      behindDocument: true,
      allowOverlap: true,
      layoutInCell: true,
    },
  });

  return new Header({
    children: [
      new Paragraph({
        spacing: P0,
        children: [watermarkImage],
      }),
    ],
  });
}

function buildFooterExactLikeImage(wmText: string) {
  const { left, right } = splitWatermarkText(wmText);

  const red = "FF6B6B";
  const blue = "3B5BFF";

  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 0 },
        children: [
          run(left ? `${left} ` : "CONFIDENTIAL REPORT: ", { color: red, size: 22 }),
          run(right || "RACE Innovations Pvt ltd.", { color: blue, underline: true, size: 22 }),
        ],
      }),
    ],
  });
}

/** =========================
 * Table helpers
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
        spacing: P0,
        children: [run(text, { bold: true, color: "FFFFFF", size: 24 })],
      }),
    ],
  });
}

function makeHeaderRow() {
  return new TableRow({
    tableHeader: true,
    cantSplit: true,
    height: { value: 711, rule: HeightRule.ATLEAST },
    children: [
      headerCell("GPS\nNO", 1),
      headerCell("KMS", 1),
      headerCell("NE\nCORDINATE", 1),
      headerCell("DETAILS", 2),
      headerCell("LOCATION", 2),
      headerCell("PHOTO", 3),
      headerCell("VEHICLE\nMOVEMENT", 1),
    ],
  });
}

function textCell(text: string, span: number, align: AlignmentType, vAlign: VerticalAlign) {
  const lines = splitLines(text);
  const paras =
    align === AlignmentType.LEFT ? lines.map(paragraphFromLine) : lines.map((ln) => paragraphPlain(ln, align));

  return new TableCell({
    columnSpan: span,
    verticalAlign: vAlign,
    borders: CELL_BORDERS,
    children: paras,
  });
}

/** =========================
 * Movement square (color only)
 * ========================= */
async function squarePngBytes(colorHex: string, sizePx = 26): Promise<Uint8Array> {
  const canvas = document.createElement("canvas");
  canvas.width = sizePx;
  canvas.height = sizePx;

  const ctx = canvas.getContext("2d")!;
  ctx.fillStyle = colorHex;
  ctx.fillRect(0, 0, sizePx, sizePx);

  ctx.strokeStyle = "#111111";
  ctx.lineWidth = 2;
  ctx.strokeRect(1, 1, sizePx - 2, sizePx - 2);

  const base64 = canvas.toDataURL("image/png").split(",")[1];
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

const MOVEMENT_SQUARE_CACHE = new Map<string, Uint8Array>();

async function movementCell(movement: string) {
  const m = normalizeMovement(movement);
  const color =
    m === "red" ? "#FF0000" : m === "yellow" ? "#FFC000" : m === "green" ? "#00B050" : "#FFFFFF";

  const box = 24;
  let bytes = MOVEMENT_SQUARE_CACHE.get(color);
  if (!bytes) {
    bytes = await squarePngBytes(color, box);
    MOVEMENT_SQUARE_CACHE.set(color, bytes);
  }

  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    borders: CELL_BORDERS,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: P0,
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
    const bmp = await createImageBitmap(blob);
    const canvas = document.createElement("canvas");
    canvas.width = bmp.width;
    canvas.height = bmp.height;

    const ctx = canvas.getContext("2d")!;
    ctx.drawImage(bmp, 0, 0);

    const pngBlob: Blob = await new Promise((resolve) => canvas.toBlob((b) => resolve(b as Blob), "image/png"));
    return new Uint8Array(await pngBlob.arrayBuffer());
  } catch {
    return null;
  }
}

async function fetchBytes(url: string, timeoutMs = DEFAULT_PHOTO_TIMEOUT_MS) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, { signal: controller.signal });
    if (!res.ok) throw new Error(`Photo fetch failed: ${res.status}`);

    const blob = await res.blob();

    if (blob.type === "image/webp") {
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
      return await blobToBytes(dl.data);
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
    const idx = u.pathname.indexOf("/storage/v1/object/");
    if (idx === -1) return null;
    const tail = u.pathname.slice(idx + "/storage/v1/object/".length);
    const parts = tail.split("/").filter(Boolean);
    if (parts.length < 3) return null;
    const bucket = parts[1];
    const path = parts.slice(2).join("/");
    if (!bucket || !path) return null;
    return { bucket, path };
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
        return await fetchBytes(raw);
      } catch {
        return null;
      }
    }

    const cleaned = cleanPath(raw);
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

async function photoCell(supabase: any, refs: string[], includePhotos: boolean) {
  const list = (refs || []).filter(Boolean).slice(0, 3);

  if (!includePhotos || list.length === 0) {
    return new TableCell({
      columnSpan: 3,
      verticalAlign: VerticalAlign.TOP,
      borders: CELL_BORDERS,
      children: [new Paragraph({ spacing: P0, text: "" })],
    });
  }

  const multi = list.length > 1;
  const imgW = 277;
  const imgH = multi ? 150 : 208;

  const bytesList = await Promise.all(list.map((r) => resolvePhotoBytes(supabase, r)));

  const paras: Paragraph[] = [];
  for (let i = 0; i < bytesList.length; i++) {
    const bytes = bytesList[i];
    if (!bytes) continue;
    paras.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { ...P0, after: i === bytesList.length - 1 ? 0 : 60 },
        children: [new ImageRun({ data: bytes, transformation: { width: imgW, height: imgH } })],
      })
    );
  }

  if (paras.length === 0 && DEBUG_PHOTOS) {
    throw new Error(`Photo refs detected but could not be resolved.\nFirst ref: ${list[0]}`);
  }

  return new TableCell({
    columnSpan: 3,
    verticalAlign: VerticalAlign.TOP,
    borders: CELL_BORDERS,
    children: paras.length ? paras : [new Paragraph({ spacing: P0, text: "" })],
  });
}

/** =========================
 * Normalize Point
 * ✅ DETAILS = category (default) OR point.details/description
 * ✅ LOCATION = readable location from reverse geocode (later)
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

  const latRaw = raw.loc_lat ?? raw.lat ?? raw.latitude ?? raw.north ?? raw.n;
  const lonRaw = raw.loc_lon ?? raw.lon ?? raw.lng ?? raw.longitude ?? raw.east ?? raw.e;

  const lat = latRaw != null ? Number(latRaw) : null;
  const lon = lonRaw != null ? Number(lonRaw) : null;

  let ne_coordinate = "";
  if (lat != null && lon != null && !Number.isNaN(lat) && !Number.isNaN(lon)) {
    ne_coordinate = `N${s(lat)}\nE${s(lon)}`;
  } else if (typeof raw.ne_coordinate === "string" && raw.ne_coordinate.trim()) {
    ne_coordinate = raw.ne_coordinate.trim();
  } else if (typeof raw.coordinate === "string" && raw.coordinate.trim()) {
    ne_coordinate = raw.coordinate.trim();
  }

  // ✅ DETAILS: use point details/remarks if present. otherwise fallback to report category
  const details = s(raw.details ?? raw.remarks ?? raw.note ?? raw.description ?? raw.desc ?? raw.__report_category ?? "");

  // ✅ LOCATION: keep only text fields if present (DO NOT put maps URL here)
  const location = s(raw.location ?? raw.place ?? raw.area ?? raw.city ?? raw.village ?? "");

  const photo_refs = Array.from(new Set(collectImageStrings(raw)));

  // ✅ MOVEMENT: use report difficulty OR row-level difficulty
  const movement = normalizeMovement(
    raw.difficulty ?? raw.vehicle_movement ?? raw.movement ?? raw.status ?? raw.__report_difficulty ?? ""
  );

  return {
    gps_no: gpsCandidate,
    kms: kmsCandidate,
    ne_coordinate,
    details,
    location,
    photo_refs,
    movement,
    __lat: !Number.isNaN(lat as any) ? lat : null,
    __lon: !Number.isNaN(lon as any) ? lon : null,
  };
}

/** =========================
 * GPS / KMS compute
 * ========================= */
function haversineKm(lat1: number, lon1: number, lat2: number, lon2: number) {
  const R = 6371;
  const toRad = (x: number) => (x * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

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
 * Extra photo table lookup (report_photos)
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
 * (tries many tables)
 * ========================= */
const TABLE_CANDIDATES = [
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
];

async function getPointsForReport(supabase: any, reportId: string) {
  const { data: report, error: rErr } = await supabase.from("reports").select("*").eq("id", reportId).single();
  if (rErr) throw rErr;

  const routeId = report?.route_id ?? report?.routeId ?? null;
  const projectId = report?.project_id ?? report?.projectId ?? null;

  // ✅ Difficulty from reports table
  const reportDifficulty = normalizeMovement(report?.difficulty ?? "");

  for (const table of TABLE_CANDIDATES) {
    try {
      // small existence check
      const probe = await supabase.from(table).select("*").limit(1);
      if (probe.error) continue;

      for (const fk of FK_CANDIDATES) {
        const targetValue =
          fk.toLowerCase().includes("route") ? routeId : fk.toLowerCase().includes("project") ? projectId : reportId;

        if (!targetValue) continue;

        const { data, error } = await supabase.from(table).select("*").eq(fk, targetValue);
        if (!error && Array.isArray(data) && data.length) {
          // ✅ patch rows with report difficulty + report category (so DETAILS can fallback)
          const patched = (data || []).map((row: any) => ({
            ...row,
            __report_difficulty: reportDifficulty,
            __report_category: report?.category ?? "",
          }));
          return { points: patched, report, routeId };
        }
      }
    } catch {}
  }

  // fallback to report location if exists
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
        },
      ],
      report,
      routeId,
    };
  }

  throw new Error(`Points not found.\nreport_id=${reportId}\nroute_id=${routeId || "NULL"} project_id=${projectId || "NULL"}`);
}

/** =========================
 * ✅ BODY ROW BUILDER
 * - LOCATION is reverse-geocoded from NE coords
 * ========================= */
async function makeBodyRow(supabase: any, p: NormalizedPoint, includePhotos: boolean) {
  // 1) Determine lat/lon
  let lat = p.__lat ?? null;
  let lon = p.__lon ?? null;

  if ((lat == null || lon == null) && p.ne_coordinate) {
    const parsed = parseNEToDecimal(p.ne_coordinate);
    if (parsed) {
      lat = parsed.lat;
      lon = parsed.lon;
    }
  }

  // 2) Reverse geocode to readable place
  let locText = (p.location || "").trim();
  if (lat != null && lon != null && Number.isFinite(lat) && Number.isFinite(lon)) {
    const place = await reverseGeocodeOSM(lat, lon);
    if (place) locText = place;
  }
  if (!locText) locText = "—";

  return new TableRow({
    cantSplit: true,
    height: { value: 2490, rule: HeightRule.ATLEAST },
    children: [
      textCell(p.gps_no, 1, AlignmentType.CENTER, VerticalAlign.CENTER),
      textCell(p.kms, 1, AlignmentType.CENTER, VerticalAlign.CENTER),
      textCell(p.ne_coordinate, 1, AlignmentType.CENTER, VerticalAlign.CENTER),
      textCell(p.details, 2, AlignmentType.LEFT, VerticalAlign.CENTER),
      textCell(locText, 2, AlignmentType.LEFT, VerticalAlign.CENTER),
      await photoCell(supabase, p.photo_refs, includePhotos),
      await movementCell(p.movement), // ✅ based on difficulty
    ],
  });
}

/** =========================
 * DOC builder
 * ========================= */
async function buildDoc(opts: {
  supabase: any;
  includePhotos: boolean;
  fileName: string;
  points: any[];
  extraPhotoRefs?: string[];
  watermark?: WatermarkOptions;
  autoSave?: boolean;
}): Promise<Blob> {
  const rows: TableRow[] = [makeHeaderRow()];

  let normalized = enrichPointsAlways(opts.points);

  if (opts.extraPhotoRefs?.length) {
    normalized = applyExtraPhotos(normalized, opts.extraPhotoRefs);
  }

  for (const p of normalized) {
    rows.push(await makeBodyRow(opts.supabase, p, opts.includePhotos));
  }

  const table = new Table({
    style: "Table Grid Light1",
    layout: TableLayoutType.FIXED,
    width: { size: TABLE_TOTAL_W, type: WidthType.DXA },
    columnWidths: GRID_COLS,
    rows,
    alignment: AlignmentType.CENTER as any,
  });

  const wmEnabled = !!opts.watermark?.enabled;
  const wmText = (opts.watermark?.text || "").trim();

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { orientation: PageOrientation.LANDSCAPE },
            margin: { left: 1440, right: 1440, top: 1418, bottom: 820, header: 720, footer: 720 } as any,
            ...(pageBordersTSPL() as any),
          } as any,
        },
        headers: wmEnabled ? { default: await buildHeaderWithDiagonalWatermark(wmText) } : undefined,
        footers: wmEnabled ? { default: buildFooterExactLikeImage(wmText) } : undefined,
        children: [table],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);

  if (opts.autoSave !== false) {
    saveAs(blob, opts.fileName);
  }

  return blob;
}

/** =========================
 * EXPORTED DOCX functions
 * ========================= */
export async function downloadReportDOCX(supabase: any, reportId: string, opts: DownloadOpts = {}) {
  const includePhotos = opts.includePhotos ?? true;

  const { points } = await getPointsForReport(supabase, reportId);
  const extraPhotoRefs = includePhotos ? await getExtraPhotosForReport(supabase, reportId) : [];

  await buildDoc({
    supabase,
    includePhotos,
    fileName: opts.fileName || `report-${String(reportId).slice(0, 8)}.docx`,
    points,
    extraPhotoRefs,
    watermark: opts.watermark,
  });
}

export async function generateProjectDOCX(
  supabase: any,
  projectId: string,
  opts: DownloadOpts = {}
): Promise<{ blob: Blob; fileName: string }> {
  const includePhotos = opts.includePhotos ?? true;

  const { data: project, error: pErr } = await supabase.from("projects").select("*").eq("id", projectId).single();
  if (pErr) throw pErr;

  const { data: reports, error: rErr } = await supabase
    .from("reports")
    .select("*")
    .eq("project_id", projectId)
    .order("created_at", { ascending: true });
  if (rErr) throw rErr;

  const allPoints: any[] = [];
  const allExtraPhotos: string[] = [];

  for (const r of (reports || []) as ReportRow[]) {
    const { points } = await getPointsForReport(supabase, r.id);
    allPoints.push(...(points || []));

    if (includePhotos) {
      const extra = await getExtraPhotosForReport(supabase, r.id);
      allExtraPhotos.push(...extra);
    }
  }

  const fileName = opts.fileName || `${projectNameOf(project as ProjectRow)}-ALL-REPORTS.docx`;

  const blob = await buildDoc({
    supabase,
    includePhotos,
    fileName,
    points: allPoints,
    extraPhotoRefs: Array.from(new Set(allExtraPhotos)),
    watermark: opts.watermark,
    autoSave: false,
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

  const { data: project, error: pErr } = await supabase.from("projects").select("*").eq("id", projectId).single();
  if (pErr) throw pErr;

  const allPoints: any[] = [];
  const allExtraPhotos: string[] = [];

  for (const reportId of reportIds) {
    const { points } = await getPointsForReport(supabase, reportId);
    allPoints.push(...(points || []));

    if (includePhotos) {
      const extra = await getExtraPhotosForReport(supabase, reportId);
      allExtraPhotos.push(...extra);
    }
  }

  const name = projectNameOf(project as ProjectRow);
  const fileName = opts.fileName || `${name}-REPORTS-${reportIds.length}.docx`;

  const blob = await buildDoc({
    supabase,
    includePhotos,
    fileName,
    points: allPoints,
    extraPhotoRefs: Array.from(new Set(allExtraPhotos)),
    watermark: opts.watermark,
    autoSave: false,
  });

  return { blob, fileName };
}

/** =========================
 * EXPORTED GPX functions
 * ========================= */
async function collectGpxPointsForReportId(supabase: any, reportId: string): Promise<GPXPoint[]> {
  const { points, report } = await getPointsForReport(supabase, reportId);
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
