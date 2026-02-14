import { NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  HeadingLevel,
  ImageRun,
  ShadingType,
  VerticalAlign,
} from "docx";

export const runtime = "nodejs"; // IMPORTANT (not edge)

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL!,
  process.env.SUPABASE_SERVICE_ROLE_KEY!,
  { auth: { persistSession: false } }
);

function headerCell(text: string) {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, color: "FFFFFF", fill: "2F5E8F" },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text, color: "FFFFFF", bold: true })],
      }),
    ],
  });
}

function normalCell(text: string) {
  return new TableCell({
    children: [new Paragraph({ children: [new TextRun(text || "")] })],
  });
}

async function fetchImageBytes(url: string): Promise<Uint8Array | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const ab = await res.arrayBuffer();
    return new Uint8Array(ab);
  } catch {
    return null;
  }
}

export async function GET(_: Request, { params }: { params: { id: string } }) {
  try {
    const reportId = params.id;

    const { data: report, error: repErr } = await supabaseAdmin
      .from("reports")
      .select("*")
      .eq("id", reportId)
      .single();

    if (repErr || !report) {
      return NextResponse.json({ error: repErr?.message || "Report not found" }, { status: 404 });
    }

    // project
    let projectName = "";
    if (report.project_id) {
      const { data: proj } = await supabaseAdmin.from("projects").select("*").eq("id", report.project_id).single();
      projectName = proj?.name || "";
    }

    // photos
    const { data: photos } = await supabaseAdmin
      .from("report_photos")
      .select("*")
      .eq("report_id", reportId)
      .order("created_at", { ascending: true });

    // points
    const { data: points } = await supabaseAdmin
      .from("report_path_points")
      .select("*")
      .eq("report_id", reportId)
      .order("seq", { ascending: true });

    const photoUrls: string[] =
      (photos || [])
        .map((p: any) => p.url)
        .filter((u: any) => typeof u === "string" && u.startsWith("http")) || [];

    // Build table rows like your sample format :contentReference[oaicite:1]{index=1}
    const header = new TableRow({
      children: [
        headerCell("GPS NO"),
        headerCell("KMS"),
        headerCell("NE COORDINATE"),
        headerCell("DETAILS"),
        headerCell("LOCATION"),
        headerCell("PHOTO"),
        headerCell("VEHICLE MOVEMENT"),
      ],
    });

    const rows: TableRow[] = [header];

    for (const pt of points || []) {
      const lat = pt.latitude ?? "";
      const lng = pt.longitude ?? "";
      const ne = lat && lng ? `N${lat} E${lng}` : "";

      // attach up to 2 photos per row (simple approach)
      // If you have photo per point mapping in DB, you can replace this logic.
      const img1 = photoUrls[0] ? await fetchImageBytes(photoUrls[0]) : null;
      const img2 = photoUrls[1] ? await fetchImageBytes(photoUrls[1]) : null;

      const photoCellChildren: Paragraph[] = [];

      if (img1) {
        photoCellChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                data: img1,
                transformation: { width: 220, height: 140 },
              }),
            ],
          })
        );
      }

      if (img2) {
        photoCellChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                data: img2,
                transformation: { width: 220, height: 140 },
              }),
            ],
          })
        );
      }

      if (!photoCellChildren.length) {
        photoCellChildren.push(new Paragraph("")); // empty
      }

      rows.push(
        new TableRow({
          children: [
            normalCell(String(pt.seq ?? "")),
            normalCell(String(pt.km ?? "")),
            normalCell(ne),
            normalCell(pt.details ?? ""),
            normalCell(pt.location_text ?? ""),
            new TableCell({ children: photoCellChildren }),
            normalCell(pt.vehicle_movement ?? ""),
          ],
        })
      );
    }

    const table = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows,
    });

    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              text: "ROUTE SURVEY REPORT",
              heading: HeadingLevel.TITLE,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: projectName ? `Project: ${projectName}` : "Project",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: `Report: ${report.category || "Report"}`,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: report.created_at ? new Date(report.created_at).toLocaleString() : "",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({ text: "" }),
            new Paragraph({
              text: "Stage Details",
              heading: HeadingLevel.HEADING_1,
            }),
            table,
          ],
        },
      ],
    });

    const buf = await Packer.toBuffer(doc);
    const base64 = buf.toString("base64");
    const filename = `${(projectName || "project").replaceAll(" ", "_")}_${(report.category || "report").replaceAll(
      " ",
      "_"
    )}.docx`;

    return NextResponse.json({ base64, filename });
  } catch (e: any) {
    return NextResponse.json({ error: e?.message || String(e) }, { status: 500 });
  }
}
