export type BackendExportPayload = {
  mode: "single-report" | "project" | "project-selected";
  reportId?: string;
  projectId?: string;
  reportIds?: string[];
  fileName?: string;
  includePhotos?: boolean;
  watermark?: { enabled?: boolean; text?: string };
  cover?: any;
};

function extractFileName(disposition: string | null, fallback: string) {
  if (!disposition) return fallback;
  const match = disposition.match(/filename="?([^";]+)"?/i);
  return match?.[1] || fallback;
}

export async function downloadDocxViaBackend(payload: BackendExportPayload) {
  const res = await fetch("/api/reports/export/project-docx", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    let message = "Failed to export DOCX.";
    try {
      const json = await res.json();
      message = json?.message || message;
    } catch {}
    throw new Error(message);
  }

  const blob = await res.blob();
  const fallback = payload.fileName || "report-export.docx";
  const fileName = extractFileName(res.headers.get("Content-Disposition"), fallback);

  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.URL.revokeObjectURL(url);
}
