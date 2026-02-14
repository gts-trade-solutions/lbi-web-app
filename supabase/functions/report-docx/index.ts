// supabase/functions/reports-docx/index.ts
import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "jsr:@supabase/supabase-js@2";

import {
  AlignmentType,
  Document,
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
  VerticalAlign,
  WidthType,
} from "npm:docx@8.5.0";

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const corsHeaders: Record<string, string> = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Expose-Headers": "Content-Disposition, X-Reports-Included",
};

const DEFAULT_MAX_REPORTS = 50;
const DEFAULT_MAX_POINTS = 400;
const DEFAULT_MAX_PHOTOS_PER_POINT = 2;
const DEFAULT_MAX_TOTAL_IMAGES_PER_REPORT = 80;
const MAX_IMAGE_BYTES = 2_000_000; // 2MB

function json(obj: any, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

function tableBorders() {
  return {
    top: { style: "single", size: 4, color: "000000" },
    bottom: { style: "single", size: 4, color: "000000" },
    left: { style: "single", size: 4, color: "000000" },
    right: { style: "single", size: 4, color: "000000" },
    insideHorizontal: { style: "single", size: 4, color: "000000" },
    insideVertical: { style: "single", size: 4, color: "000000" },
  } as const;
}

function toNumber(v: any): number | null {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string") {
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function toDMM(lat: number, lon: number) {
  const fmt = (v: number, pos: string, neg: string, pad: number) => {
    const hemi = v >= 0 ? pos : neg;
    const abs = Math.abs(v);
    const deg = Math.floor(abs);
    const min = (abs - deg) * 60;
    const degStr = String(deg).padStart(pad, "0");
    const minStr = min.toFixed(3).padStart(6, "0");
    return `${hemi}${degStr} ${minStr}`;
  };
  return `${fmt(lat, "N", "S", 2)} ${fmt(lon, "E", "W", 3)}`;
}

function pick(obj: any, keys: string[], fallback = "") {
  for (const k of keys) {
    const v = obj?.[k];
    if (v !== undefined && v !== null && v !== "") return v;
  }
  return fallback;
}

async function fetchImageBytes(url: string): Promise<Uint8Array | null> {
  try {
    const r = await fetch(url);
    if (!r.ok) return null;
    const ab = await r.arrayBuffer();
    if (ab.byteLength > MAX_IMAGE_BYTES) return null;
    return new Uint8Array(ab);
  } catch {
    return null;
  }
}

async function downloadFromStorage(
  supabase: any,
  bucket: string,
  path: string
): Promise<Uint8Array | null> {
  try {
    const { data, error } = await supabase.storage.from(bucket).download(path);
    if (error || !data) return null;
    const ab = await data.arrayBuffer();
    if (ab.byteLength > MAX_IMAGE_BYTES) return null;
    return new Uint8Array(ab);
  } catch {
    return null;
  }
}

function mkHeaderCell(label: string) {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "2F5597" },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: label, bold: true, color: "FFFFFF" })],
      }),
    ],
  });
}

function mkCellCenter(text: string) {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: AlignmentType.CENTER, text: text ?? "" })],
  });
}

function mkCellLeft(text: string) {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: AlignmentType.LEFT, text: text ?? "" })],
  });
}

function mkPhotoCell(imageRuns: ImageRun[]) {
  const children: Paragraph[] = [];
  if (!imageRuns?.length) children.push(new Paragraph(""));
  else {
    for (let i = 0; i < imageRuns.length; i++) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [imageRuns[i]],
        })
      );
      if (i !== imageRuns.length - 1) children.push(new Paragraph(""));
    }
  }
  return new TableCell({ verticalAlign: VerticalAlign.CENTER, children });
}

function titleCenter(text: string) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, bold: true, size: 30, color: "98A2B3" })],
    spacing: { after: 160 },
  });
}

function headingBlue(text: string) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, color: "2F5597", size: 32 })],
    spacing: { after: 200 },
  });
}

function smallMuted(text: string) {
  return new Paragraph({
    children: [new TextRun({ text, color: "667085", size: 20 })],
    spacing: { after: 120 },
  });
}

async function loadPoints(supabase: any, reportId: string, maxPoints: number) {
  const tryTables = ["report_path_points", "report_points", "path_points"];
  for (const t of tryTables) {
    const { data, error } = await supabase
      .from(t)
      .select("*")
      .eq("report_id", reportId)
      .order("seq", { ascending: true })
      .limit(maxPoints);

    if (!error && Array.isArray(data)) return data as any[];
  }
  return [];
}

async function loadPhotos(supabase: any, reportId: string) {
  const tryTables = ["report_photos", "photos"];
  for (const t of tryTables) {
    const { data, error } = await supabase
      .from(t)
      .select("*")
      .eq("report_id", reportId)
      .order("created_at", { ascending: true });

    if (!error && Array.isArray(data)) return data as any[];
  }
  return [];
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    const body = await req.json().catch(() => ({}));

    const reportIds = Array.isArray(body?.reportIds) ? body.reportIds : [];
    const includePhotos = body?.includePhotos !== false;

    const maxReports = Number(body?.maxReports ?? DEFAULT_MAX_REPORTS);
    const maxPoints = Number(body?.maxPoints ?? DEFAULT_MAX_POINTS);
    const maxPhotosPerPoint = Number(body?.maxPhotosPerPoint ?? DEFAULT_MAX_PHOTOS_PER_POINT);
    const maxTotalImagesPerReport = Number(body?.maxTotalImagesPerReport ?? DEFAULT_MAX_TOTAL_IMAGES_PER_REPORT);

    if (!reportIds.length) return json({ error: "reportIds[] required" }, 400);

    const supabaseUrl = Deno.env.get("SUPABASE_URL") || Deno.env.get("PROJECT_URL");
    const serviceKey = Deno.env.get("SERVICE_ROLE_KEY");
    if (!supabaseUrl || !serviceKey) {
      return json({ error: "Missing SUPABASE_URL/PROJECT_URL or SERVICE_ROLE_KEY" }, 500);
    }

    const supabase = createClient(supabaseUrl, serviceKey, {
      auth: { persistSession: false },
    });

    const limitedReportIds = reportIds.slice(0, maxReports);

    // Load all report rows in one query
    const { data: reportsData, error: repErr } = await supabase
      .from("reports")
      .select("*")
      .in("id", limitedReportIds);

    if (repErr) throw repErr;

    const reports = (reportsData || []) as any[];

    // IMPORTANT: keep the same order as UI (reportIds order)
    const map = new Map<string, any>();
    for (const r of reports) map.set(String(r.id), r);
    const orderedReports = limitedReportIds.map((id: string) => map.get(String(id))).filter(Boolean);

    const sections: { properties: any; children: (Paragraph | Table)[] }[] = [];

    for (const report of orderedReports) {
      const reportId = String(report.id);
      const projectId = report.project_id ? String(report.project_id) : null;

      // project title (if project_id null, still works)
      let projectName = "Project";
      if (projectId) {
        const { data: proj } = await supabase
          .from("projects")
          .select("name")
          .eq("id", projectId)
          .maybeSingle();
        projectName = String(proj?.name || "Project");
      }

      const reportLabel = String(report?.category || "Report");
      const created = report?.created_at ? new Date(report.created_at).toLocaleString() : "";

      const children: (Paragraph | Table)[] = [];
      children.push(titleCenter(projectName));
      children.push(smallMuted(`Report: ${reportLabel}  •  ${created}`));
      children.push(headingBlue("STAGE SUMMARY:"));
      children.push(
        new Paragraph({
          children: [new TextRun({ text: "LEGENDS:", bold: true, underline: {} })],
        })
      );
      children.push(new Paragraph("Obstruction has to be cleared for safe movement"));
      children.push(new Paragraph("Critical, Cargo movement is impossible"));
      children.push(new Paragraph("Not critical"));
      children.push(new Paragraph(""));

      // points
      let points = await loadPoints(supabase, reportId, maxPoints);

      // fallback to a single row if no points table exists
      if (points.length === 0) {
        const lat = toNumber(report.loc_lat);
        const lon = toNumber(report.loc_lon);
        points = [
          {
            seq: 1,
            kms: report.kms ?? report.km ?? 0,
            latitude: lat,
            longitude: lon,
            details: report.category || "",
            remarks: report.description || "",
            location: report.location || report.address || "",
            vehicle_movement: report.vehicle_movement || "",
          },
        ];
      }

      const photos = await loadPhotos(supabase, reportId);

      const photosBySeq = new Map<number, any[]>();
      const loose: any[] = [];

      for (const p of photos) {
        const rawKey =
          p.point_seq ?? p.path_point_seq ?? p.seq ?? p.gps_no ?? p.point_no ?? null;
        const keyNum = rawKey === null || rawKey === undefined ? NaN : Number(rawKey);
        if (Number.isFinite(keyNum)) {
          if (!photosBySeq.has(keyNum)) photosBySeq.set(keyNum, []);
          photosBySeq.get(keyNum)!.push(p);
        } else {
          loose.push(p);
        }
      }

      const rows: TableRow[] = [];
      rows.push(
        new TableRow({
          children: [
            mkHeaderCell("GPS NO"),
            mkHeaderCell("KMS"),
            mkHeaderCell("NE COORDINATE"),
            mkHeaderCell("DETAILS"),
            mkHeaderCell("LOCATION"),
            mkHeaderCell("PHOTO"),
            mkHeaderCell(""),
          ],
        })
      );

      let gpsNo = 1;
      let looseIdx = 0;
      let totalImagesUsed = 0;

      for (let i = 0; i < points.length; i++) {
        const pt = points[i];

        const kms = String(pick(pt, ["kms", "km", "distance_km", "distance"], ""));
        const lat = toNumber(pick(pt, ["latitude", "lat"], null));
        const lon = toNumber(pick(pt, ["longitude", "lon", "lng"], null));
        const coord = lat !== null && lon !== null ? toDMM(lat, lon) : "—";

        const details =
          String(pick(pt, ["details", "note", "description"], pick(report, ["category"], ""))) +
          (pick(pt, ["remarks"], "") ? `\n${pick(pt, ["remarks"], "")}` : "");

        const location = String(
          pick(pt, ["location", "address"], pick(report, ["location", "address"], ""))
        );

        const seq = Number(pick(pt, ["seq", "gps_no"], i + 1));

        let pointPhotos = photosBySeq.get(seq) || [];
        if (pointPhotos.length === 0 && looseIdx < loose.length) {
          pointPhotos = loose.slice(looseIdx, looseIdx + maxPhotosPerPoint);
          looseIdx += pointPhotos.length;
        }
        pointPhotos = pointPhotos.slice(0, maxPhotosPerPoint);

        const imageRuns: ImageRun[] = [];

        if (includePhotos) {
          for (const p of pointPhotos) {
            if (totalImagesUsed >= maxTotalImagesPerReport) break;

            let bytes: Uint8Array | null = null;
            if (p?.bucket && p?.path) {
              bytes = await downloadFromStorage(supabase, p.bucket, p.path);
            } else {
              const u = p?.url || p?.public_url || "";
              if (u) bytes = await fetchImageBytes(u);
            }

            if (bytes) {
              totalImagesUsed++;
              imageRuns.push(
                new ImageRun({
                  data: bytes,
                  transformation: { width: 340, height: 220 },
                })
              );
            }
          }
        }

        const movementText = String(
          pick(pt, ["vehicle_movement", "movement", "status"], "")
        ).toLowerCase();

        let fill = "00B050";
        if (movementText.includes("critical") || movementText.includes("not possible")) fill = "FF0000";
        else if (movementText.includes("slow") || movementText.includes("careful")) fill = "FFFF00";

        const indicatorCell = new TableCell({
          verticalAlign: VerticalAlign.CENTER,
          shading: { type: ShadingType.CLEAR, color: "auto", fill },
          children: [new Paragraph("")],
        });

        rows.push(
          new TableRow({
            children: [
              mkCellCenter(String(gpsNo++)),
              mkCellCenter(kms),
              mkCellCenter(coord),
              mkCellLeft(details),
              mkCellLeft(location),
              mkPhotoCell(imageRuns),
              indicatorCell,
            ],
          })
        );
      }

      const table = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        borders: tableBorders(),
        rows,
      });

      children.push(table);

      sections.push({
        properties: {
          page: {
            size: { orientation: PageOrientation.LANDSCAPE },
            margin: { top: 720, right: 500, bottom: 720, left: 500 },
          },
        },
        children,
      });
    }

    const doc = new Document({ sections });
    const buf = await Packer.toBuffer(doc);

    const filename = `overall-reports-${Date.now()}.docx`;

    return new Response(buf, {
      status: 200,
      headers: {
        ...corsHeaders,
        "Content-Type": DOCX_MIME,
        "Content-Disposition": `attachment; filename="${filename}"`,
        "X-Reports-Included": String(sections.length),
      },
    });
  } catch (e: any) {
    return json({ error: e?.message || String(e) }, 500);
  }
});
