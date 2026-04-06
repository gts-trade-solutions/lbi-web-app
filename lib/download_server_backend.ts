/**
 * lib/download_server_backend.ts
 * Server-side wrapper for API routes
 */

import {
  generateProjectDOCX,
  generateProjectDOCXByReportIds,
  generateProjectGPX,
  generateProjectGPXByReportIds,
} from "./download_core";

export async function generateProjectDOCXBuffer(
  supabase: any,
  projectId: string,
  opts: any = {}
) {
  const { blob, fileName } = await generateProjectDOCX(supabase, projectId, opts);
  const arrayBuffer = await blob.arrayBuffer();
  return {
    buffer: Buffer.from(arrayBuffer),
    fileName,
  };
}

export async function generateProjectDOCXByReportIdsBuffer(
  supabase: any,
  projectId: string,
  reportIds: string[],
  opts: any = {}
) {
  const { blob, fileName } = await generateProjectDOCXByReportIds(
    supabase,
    projectId,
    reportIds,
    opts
  );
  const arrayBuffer = await blob.arrayBuffer();
  return {
    buffer: Buffer.from(arrayBuffer),
    fileName,
  };
}

export {
  generateProjectGPX,
  generateProjectGPXByReportIds,
  generateProjectDOCX,
  generateProjectDOCXByReportIds,
};