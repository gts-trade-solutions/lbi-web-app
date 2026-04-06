import { NextRequest, NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";
import { generateProjectDOCXByReportIdsBuffer } from "../../../lib/report_docx_server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Body = {
  projectId?: string;
  reportIds?: string[];
  fileName?: string;
  includePhotos?: boolean;
  watermark?: { enabled?: boolean; text?: string };
};

function jsonError(message: string, status = 400) {
  return NextResponse.json({ success: false, message }, { status });
}

export async function POST(req: NextRequest) {
  try {
    const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
    const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

    if (!supabaseUrl || !supabaseAnonKey) {
      return jsonError("Supabase environment variables are missing.", 500);
    }

    const body = (await req.json()) as Body;

    const projectId = String(body?.projectId || "").trim();
    const reportIds = Array.isArray(body?.reportIds)
      ? body.reportIds.map((x) => String(x || "").trim()).filter(Boolean)
      : [];
    const fileName = String(body?.fileName || "export.docx").trim() || "export.docx";
    const includePhotos = body?.includePhotos ?? true;
    const watermark = body?.watermark ?? { enabled: false, text: "" };

    if (!projectId) return jsonError("projectId is required.");
    if (!reportIds.length) return jsonError("At least 1 reportId is required.");

    const authHeader = req.headers.get("authorization") || "";

    const supabase = createClient(supabaseUrl, supabaseAnonKey, {
      global: {
        headers: authHeader ? { Authorization: authHeader } : {},
      },
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    });

    if (authHeader) {
      const token = authHeader.replace(/^Bearer\s+/i, "").trim();
      if (token) {
        const { data, error } = await supabase.auth.getUser(token);
        if (error || !data?.user) {
          return jsonError("Unauthorized export request.", 401);
        }
      }
    }

    const { buffer, fileName: resolvedName, contentType } =
      await generateProjectDOCXByReportIdsBuffer(
        supabase,
        projectId,
        reportIds,
        {
          includePhotos,
          fileName,
          watermark,
        }
      );

    return new NextResponse(buffer, {
      status: 200,
      headers: {
        "Content-Type": contentType,
        "Content-Disposition": `attachment; filename="${resolvedName.replace(/"/g, "")}"`,
        "Cache-Control": "no-store, max-age=0",
      },
    });
  } catch (e: any) {
    return jsonError(e?.message || "Export failed.", 500);
  }
}