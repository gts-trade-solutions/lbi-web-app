export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function cleanUrl(url: string) {
  return url.replace(/\/+$/, "");
}

// ✅ TEST ROUTE (open in browser)
export async function GET() {
  return new Response("OK: /api/projects-docx route is working", {
    status: 200,
    headers: { "Content-Type": "text/plain" },
  });
}

export async function POST(req: Request) {
  try {
    const SUPABASE_URL = process.env.NEXT_PUBLIC_SUPABASE_URL;
    const ANON_KEY = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

    if (!SUPABASE_URL)
      return new Response("NEXT_PUBLIC_SUPABASE_URL missing", { status: 500 });
    if (!ANON_KEY)
      return new Response("NEXT_PUBLIC_SUPABASE_ANON_KEY missing", { status: 500 });

    const authHeader = req.headers.get("authorization") || "";
    if (!authHeader.toLowerCase().startsWith("bearer ")) {
      return new Response("Missing Authorization Bearer token", { status: 401 });
    }

    const body = await req.json().catch(() => ({}));
    const projectIds = body?.projectIds;

    if (!Array.isArray(projectIds) || projectIds.length === 0) {
      return new Response("projectIds[] required", { status: 400 });
    }

    const endpoint = `${cleanUrl(SUPABASE_URL)}/functions/v1/projects-docx`;

    // ✅ Server-to-server call (no browser CORS issues)
    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        apikey: ANON_KEY,
        Authorization: authHeader, // forward user token
      },
      body: JSON.stringify({ projectIds }),
    });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      return new Response(txt || `Edge Function error ${res.status}`, {
        status: res.status,
        headers: { "Content-Type": "text/plain" },
      });
    }

    const ab = await res.arrayBuffer();

    return new Response(ab, {
      status: 200,
      headers: {
        "Content-Type": DOCX_MIME,
        "Content-Disposition": `attachment; filename="projects-route-report.docx"`,
      },
    });
  } catch (e: any) {
    return new Response(String(e?.message || e), { status: 500 });
  }
}
