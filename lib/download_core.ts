/**
 * lib/download_core.ts
 * Server-safe DOCX/GPX export core
 *
 * IMPORTANT:
 * - No "use client"
 * - No file-saver import
 */

import JSZip from "jszip";
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  HeightRule,
  ImageRun,
  Packer,
  PageNumber,
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

// =====================================================
// COPY EVERYTHING FROM YOUR CURRENT lib/download.ts
// EXCEPT:
// 1. remove: "use client";
// 2. remove: import { saveAs } from "file-saver";
// 3. keep all helper functions
// 4. keep all DOCX/GPX generator exports
// =====================================================

// You must keep/export these functions exactly:
export async function generateProjectDOCX(
  supabase: any,
  projectId: string,
  opts: any = {}
) {
  // paste your existing implementation from download.ts
  throw new Error("Paste existing implementation here");
}

export async function generateProjectDOCXByReportIds(
  supabase: any,
  projectId: string,
  reportIds: string[],
  opts: any = {}
) {
  // paste your existing implementation from download.ts
  throw new Error("Paste existing implementation here");
}

export async function generateProjectGPX(
  supabase: any,
  projectId: string,
  opts: any = {}
) {
  // paste your existing implementation from download.ts
  throw new Error("Paste existing implementation here");
}

export async function generateProjectGPXByReportIds(
  supabase: any,
  projectId: string,
  reportIds: string[],
  opts: any = {}
) {
  // paste your existing implementation from download.ts
  throw new Error("Paste existing implementation here");
}