import { supabase } from "@/lib/supabaseBrowser";
import { downloadBase64AsFile } from "@/utils/downloadBase64";

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export async function exportReportDocx(reportId: string) {
  const { data, error } = await supabase.functions.invoke("report-docx", {
    body: { reportId },
  });

  if (error) {
    // this will show real error (500/401/CORS etc.)
    throw new Error(error.message);
  }

  if (!data?.base64) {
    throw new Error("Edge function did not return { base64 }");
  }

  downloadBase64AsFile(
    data.base64,
    data.filename || `report-${reportId}.docx`,
    DOCX_MIME
  );
}
