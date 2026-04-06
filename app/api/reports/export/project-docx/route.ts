import { NextRequest, NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";
import { generateProjectDOCXByReportIdsBuffer } from "../../../../../lib/report_docx_server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Body = {
  projectId?: string;
  reportIds?: string[];
  fileName?: string;
  includePhotos?: boolean;
  watermark?: { enabled?: boolean; text?: string };
};

function getSupabaseAdmin() {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

  if (!url || !key) {
    throw new Error("Missing Supabase server environment variables.");
  }

  return createClient(url, key, {
    auth: { persistSession: false, autoRefreshToken: false },
  });
}

export async function POST(req: NextRequest) {
  try {
    const body = (await req.json()) as Body;
    const projectId = String(body?.projectId || "").trim();
    const reportIds = Array.isArray(body?.reportIds)
      ? body!.reportIds.map((x) => String(x || "").trim()).filter(Boolean)
      : [];

    if (!projectId) {
      return NextResponse.json({ success: false, message: "projectId is required." }, { status: 400 });
    }
    if (!reportIds.length) {
      return NextResponse.json({ success: false, message: "At least one reportId is required." }, { status: 400 });
    }

    const supabase = getSupabaseAdmin();
    const { buffer, fileName } = await generateProjectDOCXByReportIdsBuffer(supabase, projectId, reportIds, {
      fileName: body?.fileName,
      includePhotos: body?.includePhotos ?? true,
      watermark: body?.watermark,
    });

    return new NextResponse(buffer, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${fileName || "project-export.docx"}"`,
        "Cache-Control": "no-store",
      },
    });
  } catch (error: any) {
    return NextResponse.json(
      { success: false, message: error?.message || "Server-side DOCX export failed." },
      { status: 500 }
    );
  }
}
