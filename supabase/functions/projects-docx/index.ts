/// <reference lib="deno.unstable" />

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import { createClient } from "npm:@supabase/supabase-js@2.49.1";
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "npm:docx@8.5.0";

type ReqBody = {
  projectIds?: string[];
  includePhotos?: boolean;
  maxReports?: number;
};

function json(status: number, obj: unknown) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "content-type": "application/json" },
  });
}

function getBearer(req: Request) {
  const h = req.headers.get("authorization") || "";
  return h.startsWith("Bearer ") ? h : null;
}

function hexNoHash(s: string) {
  return (s || "").replace("#", "").toUpperCase();
}

function cellShading(fill: string) {
  return { fill: hexNoHash(fill) };
}

const BORDER_COLOR = "BFBFBF";
const HEADER_FILL = "365F91";

function bordersAll() {
  const b = { style: BorderStyle.SINGLE, size: 6, color: BORDER_COLOR };
  return {
    top: b,
    bottom: b,
    left: b,
    right: b,
    insideHorizontal: b,
    insideVertical: b,
  };
}

// ====== TABLE COLUMN SYSTEM (matches your TSPL template) ======
// 11 base columns; we "merge" them into 7 logical columns:
// [GPS][KMS][NE][DETAILS x2][LOCATION x2][PHOTO x3][MOVEMENT]
const COL_W = [950, 1150, 2070, 2960, 2960, 2240, 2240, 1930, 1930, 1930, 1410];

function headerText(txt: string) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({
        text: txt,
        bold: true,
        color: "FFFFFF",
        size: 24, // 12pt
      }),
    ],
  });
}

function normalText(txt: string, align: AlignmentType = AlignmentType.LEFT) {
  return new Paragraph({
    alignment: align,
    children: [new TextRun({ text: txt ?? "", size: 22 })], // 11pt-ish
  });
}

function emptyPara() {
  return new Paragraph({ children: [new TextRun({ text: "" })] });
}

async function blobToUint8(blob: Blob) {
  const ab = await blob.arrayBuffer();
  return new Uint8Array(ab);
}

async function getPhotosForReport(
  supaService: any,
  bucket: string | null,
  report: any
): Promise<Uint8Array[]> {
  // 1) If your reports table already has photo URLs array
  const urls: string[] =
    (Array.isArray(report.photo_urls) && report.photo_urls) ||
    (Array.isArray(report.photos) && report.photos) ||
    (report.photo_url ? [report.photo_url] : []);

  const out: Uint8Array[] = [];

  // Try URL fetch (public/signed urls)
  for (const u of urls.slice(0, 6)) {
    try {
      const r = await fetch(u);
      if (!r.ok) continue;
      const b = await r.blob();
      out.push(await blobToUint8(b));
    } catch {
      // ignore
    }
  }
  if (out.length > 0) return out;

  // 2) Else try storage folder: <reportId>/
  if (!bucket) return out;

  const { data: files } = await supaService.storage
    .from(bucket)
    .list(report.id, { limit: 6, sortBy: { column: "name", order: "asc" } });

  for (const f of files || []) {
    try {
      const { data } = await supaService.storage
        .from(bucket)
        .download(`${report.id}/${f.name}`);
      if (!data) continue;
      out.push(await blobToUint8(data));
    } catch {
      // ignore
    }
  }
  return out;
}

function buildRow(
  i: number,
  report: any,
  photos: Uint8Array[]
): TableRow {
  const gpsNo = String(i + 1);
  const kms = report.kms ?? report.km ?? report.distance_km ?? "";
  const ne =
    report.ne_coordinate ||
    report.ne ||
    (report.loc_lat != null && report.loc_lon != null
      ? `N${Number(report.loc_lat).toFixed(6)}\nE${Number(report.loc_lon).toFixed(6)}`
      : "");

  const details = report.details || report.description || "";
  const location = report.location || report.loc_address || report.place_name || "";

  // movement color
  // (adjust mapping to your DB values if needed)
  const mv = (report.vehicle_movement || report.movement || "").toString().toLowerCase();
  const mvColor =
    mv.includes("critical") ? "FF0000" : mv.includes("not") ? "FFFF00" : mv ? "00B050" : "FFFFFF";

  // photo sizing: stacked
  const count = Math.min(photos.length, 3);
  const photoParas: Paragraph[] = [];
  const w = 360;
  const h = count <= 1 ? 220 : count === 2 ? 170 : 140;

  for (const p of photos.slice(0, 3)) {
    photoParas.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60, before: 60 },
        children: [
          new ImageRun({
            data: p,
            transformation: { width: w, height: h },
          }),
        ],
      })
    );
  }
  if (photoParas.length === 0) photoParas.push(emptyPara());

  return new TableRow({
    children: [
      // GPS
      new TableCell({
        width: { size: COL_W[0], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [normalText(gpsNo, AlignmentType.CENTER)],
      }),

      // KMS
      new TableCell({
        width: { size: COL_W[1], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [normalText(String(kms), AlignmentType.CENTER)],
      }),

      // NE
      new TableCell({
        width: { size: COL_W[2], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [normalText(String(ne), AlignmentType.CENTER)],
      }),

      // DETAILS (span 2 cols)
      new TableCell({
        columnSpan: 2,
        width: { size: COL_W[3] + COL_W[4], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [normalText(String(details))],
      }),

      // LOCATION (span 2 cols)
      new TableCell({
        columnSpan: 2,
        width: { size: COL_W[5] + COL_W[6], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [normalText(String(location))],
      }),

      // PHOTO (span 3 cols)
      new TableCell({
        columnSpan: 3,
        width: { size: COL_W[7] + COL_W[8] + COL_W[9], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: photoParas,
      }),

      // MOVEMENT (color fill)
      new TableCell({
        width: { size: COL_W[10], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        shading: cellShading(mvColor),
        children: [emptyPara()],
      }),
    ],
  });
}

function buildHeaderRow(): TableRow {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: COL_W[0], type: WidthType.DXA },
        shading: cellShading(HEADER_FILL),
        verticalAlign: VerticalAlign.CENTER,
        children: [headerText("GPS\nNO")],
      }),
      new TableCell({
        width: { size: COL_W[1], type: WidthType.DXA },
        shading: cellShading(HEADER_FILL),
        verticalAlign: VerticalAlign.CENTER,
        children: [headerText("KMS")],
      }),
      new TableCell({
        width: { size: COL_W[2], type: WidthType.DXA },
        shading: cellShading(HEADER_FILL),
        verticalAlign: VerticalAlign.CENTER,
        children: [headerText("NE\nCOORDINATE")],
      }),
      new TableCell({
        columnSpan: 2,
        width: { size: COL_W[3] + COL_W[4], type: WidthType.DXA },
        shading: cellShading(HEADER_FILL),
        verticalAlign: VerticalAlign.CENTER,
        children: [headerText("DETAILS")],
      }),
      new TableCell({
        columnSpan: 2,
        width: { size: COL_W[5] + COL_W[6], type: WidthType.DXA },
        shading: cellShading(HEADER_FILL),
        verticalAlign: VerticalAlign.CENTER,
        children: [headerText("LOCATION")],
      }),
      new TableCell({
        columnSpan: 3,
        width: { size: COL_W[7] + COL_W[8] + COL_W[9], type: WidthType.DXA },
        shading: cellShading(HEADER_FILL),
        verticalAlign: VerticalAlign.CENTER,
        children: [headerText("PHOTO")],
      }),
      new TableCell({
        width: { size: COL_W[10], type: WidthType.DXA },
        shading: cellShading(HEADER_FILL),
        verticalAlign: VerticalAlign.CENTER,
        children: [headerText("VEHICLE\nMOVEMENT")],
      }),
    ],
  });
}

serve(async (req) => {
  try {
    if (req.method !== "POST") return json(405, { error: "POST only" });

    const bearer = getBearer(req);
    if (!bearer) return json(401, { error: "Missing Authorization" });

    const body = (await req.json().catch(() => ({}))) as ReqBody;

    if (!Array.isArray(body.projectIds) || body.projectIds.length === 0) {
      return json(400, { error: "projectIds[] required" });
    }

    const includePhotos = body.includePhotos !== false;
    const maxReports = Math.min(Math.max(body.maxReports ?? 50, 1), 200);

    const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
    const ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY")!;
    const SERVICE_ROLE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
    const BUCKET = Deno.env.get("REPORT_PHOTOS_BUCKET") || null;

    // auth check (user token)
    const supaAuth = createClient(SUPABASE_URL, ANON_KEY, {
      global: { headers: { Authorization: bearer } },
    });

    const { data: userData, error: userErr } = await supaAuth.auth.getUser();
    if (userErr || !userData?.user) return json(401, { error: "Invalid token" });

    // service client for DB + Storage
    const supaService = createClient(SUPABASE_URL, SERVICE_ROLE);

    // fetch project(s)
    const { data: projects } = await supaService
      .from("projects")
      .select("id,name,title,project_name,created_at")
      .in("id", body.projectIds);

    const projectName =
      projects?.[0]?.name || projects?.[0]?.title || projects?.[0]?.project_name || "Project";

    // fetch ALL reports for those projectIds
    const { data: reports, error: repErr } = await supaService
      .from("reports")
      .select("*")
      .in("project_id", body.projectIds)
      .order("created_at", { ascending: true })
      .limit(maxReports);

    if (repErr) return json(500, { error: repErr.message });
    const rows = reports || [];

    // build doc
    const docChildren: (Paragraph | Table)[] = [];

    docChildren.push(
      new Paragraph({
        text: projectName,
        alignment: AlignmentType.CENTER,
        heading: HeadingLevel.HEADING_2,
      }),
      new Paragraph({ text: "" }),
      new Paragraph({
        children: [new TextRun({ text: "STAGE SUMMARY:", bold: true, color: "1F4E79", size: 32 })],
      }),
      new Paragraph({
        children: [new TextRun({ text: "LEGENDS:", bold: true })],
      }),
      new Paragraph({ children: [new TextRun({ text: "Obstruction has to be cleared for safe movement" })] }),
      new Paragraph({ children: [new TextRun({ text: "Critical, Cargo movement is impossible" })] }),
      new Paragraph({ children: [new TextRun({ text: "Not critical" })] }),
      new Paragraph({ text: "" }),
    );

    // build rows with photos
    const tableRows: TableRow[] = [];
    tableRows.push(buildHeaderRow());

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const photos = includePhotos ? await getPhotosForReport(supaService, BUCKET, r) : [];
      tableRows.push(buildRow(i, r, photos));
    }

    const table = new Table({
      layout: TableLayoutType.FIXED,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: bordersAll(),
      columnWidths: COL_W,
      rows: tableRows,
    });

    docChildren.push(table);

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 720, bottom: 720, left: 720, right: 720 }, // 0.5"
            },
          },
          children: docChildren,
        },
      ],
    });

    const buf = await Packer.toBuffer(doc);

    const filename = `project-${body.projectIds[0]}.docx`;

    return new Response(buf, {
      status: 200,
      headers: {
        "content-type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "content-disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (e) {
    return json(500, { error: String(e?.message || e) });
  }
});
