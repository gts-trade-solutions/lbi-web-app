"use client";

import React, { useMemo, useState } from "react";
import { generateProjectDOCX, generateProjectDOCXByReportIds } from "@/lib/download";

type Props = {
  supabase: any;
  projectId: string;
  reportIds?: string[];
  className?: string;
  includePhotos?: boolean;
  buttonText?: string;
  onStart?: () => void;
  onDone?: () => void;
  onError?: (message: string) => void;
};

function normalizeIds(ids?: string[]) {
  return Array.from(new Set((ids || []).map((x) => String(x || "").trim()).filter(Boolean)));
}

export default function ProjectDocxExportClient(props: Props) {
  const {
    supabase,
    projectId,
    reportIds,
    className,
    includePhotos = true,
    buttonText = "Download DOCX",
    onStart,
    onDone,
    onError,
  } = props;

  const [loading, setLoading] = useState(false);
  const cleanReportIds = useMemo(() => normalizeIds(reportIds), [reportIds]);

  async function handleExport() {
    if (!supabase) {
      const msg = "Supabase client is missing.";
      onError?.(msg);
      alert(msg);
      return;
    }

    if (!projectId?.trim()) {
      const msg = "Project ID is missing.";
      onError?.(msg);
      alert(msg);
      return;
    }

    try {
      setLoading(true);
      onStart?.();

      if (cleanReportIds.length > 0) {
        await generateProjectDOCXByReportIds(supabase, projectId, cleanReportIds, {
          includePhotos,
        });
      } else {
        await generateProjectDOCX(supabase, projectId, {
          includePhotos,
        });
      }

      onDone?.();
    } catch (error: any) {
      const msg = error?.message || "DOCX export failed.";
      console.error("DOCX export failed:", error);
      onError?.(msg);
      alert(msg);
    } finally {
      setLoading(false);
    }
  }

  return (
    <button
      type="button"
      onClick={handleExport}
      disabled={loading}
      className={className}
    >
      {loading ? "Preparing DOCX..." : buttonText}
    </button>
  );
}
