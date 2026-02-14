"use client";

import React, { useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import { useParams } from "next/navigation";

import { supabase } from "../../../lib/supabase";
import {
  generateProjectDOCX,
  generateProjectDOCXByReportIds,
  generateProjectGPX,
  generateProjectGPXByReportIds,
} from "../../../lib/download";

type VehicleMovement = "green" | "yellow" | "red" | "";
type VMFilter = "all" | "green" | "yellow" | "red" | "unset";

type ExportFormat = "docx" | "gpx";
type ExportMode = "listed" | "selectedOne" | "selectedSplit" | "all";

type ProjectRow = {
  id: string;
  name?: string | null;
  title?: string | null;
  project_name?: string | null;
  created_at?: string | null;
};

type ReportRow = {
  id: string;
  project_id: string;
  category?: string | null;
  description?: string | null;
  created_at: string;

  // ✅ DB column is difficulty (NOT vehicle_movement)
  difficulty?: VehicleMovement | null;
};

type WatermarkOpts = { enabled: boolean; text: string };
type PreparedFile = { fileName: string; blob: Blob };

function projectNameOf(p: ProjectRow | null) {
  return p?.name || p?.title || p?.project_name || "Project";
}

function sanitizeFileBaseName(name: string) {
  const cleaned = String(name || "")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");
  return cleaned.slice(0, 80) || "Export";
}

function displayDescription(raw: string) {
  const t = (raw || "").trim();
  if (!t) return "";
  return t.replace(/^considered\s*/i, "").trim();
}

function normalizeVM(v: any): VehicleMovement {
  const t = String(v ?? "").trim().toLowerCase();
  if (t === "green") return "green";
  if (t === "yellow" || t === "amber") return "yellow";
  if (t === "red") return "red";
  return "";
}

function vmDisplayToDb(v: string): VehicleMovement {
  const t = String(v || "").trim().toLowerCase();
  if (t === "green") return "green";
  if (t === "yellow") return "yellow";
  if (t === "red") return "red";
  return "";
}

async function updateReportVM(reportId: string, next: VehicleMovement) {
  // ✅ update DB column difficulty
  const payload: any = { difficulty: next ? next : null };
  const { error } = await supabase.from("reports").update(payload).eq("id", reportId);
  if (error) throw error;
}

function vmFilterLabel(f: VMFilter) {
  if (f === "all") return "ALL";
  if (f === "unset") return "UNSET";
  return f.toUpperCase();
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function clamp(n: number, a: number, b: number) {
  return Math.max(a, Math.min(b, n));
}

function readAvgSecondsPerReport(includePhotos: boolean) {
  const key = includePhotos ? "docx_avg_sec_photo" : "docx_avg_sec_nophoto";
  const v = Number(localStorage.getItem(key) || "");
  return Number.isFinite(v) && v > 0 ? v : includePhotos ? 1.4 : 0.6;
}

function estimateSeconds(mode: ExportMode, count: number, includePhotos: boolean) {
  const per = readAvgSecondsPerReport(includePhotos);
  const base =
    mode === "all"
      ? 6
      : mode === "listed"
        ? 4
        : mode === "selectedOne"
          ? 3
          : mode === "selectedSplit"
            ? 4
            : 4;

  const est = Math.round(base + count * per);
  return clamp(est, 6, 10 * 60);
}

function parseStageRanges(input: string, total: number) {
  const out: Array<{ from: number; to: number; label: string }> = [];
  const parts = String(input || "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);

  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    const m = p.match(/^(\d+)\s*-\s*(\d+)$/) || p.match(/^(\d+)$/);
    if (!m) continue;

    let from = Number(m[1]);
    let to = Number(m[2] ?? m[1]);

    if (!Number.isFinite(from) || !Number.isFinite(to)) continue;
    if (from < 1) from = 1;
    if (to < 1) to = 1;
    if (from > total) from = total;
    if (to > total) to = total;
    if (to < from) [from, to] = [to, from];

    const label = letters[i] || `S${i + 1}`;
    out.push({ from, to, label });
  }

  return out;
}

export default function ProjectReportsPage() {
  const params = useParams<{ id: string }>();
  const projectId = params?.id;

  const [project, setProject] = useState<ProjectRow | null>(null);
  const [reports, setReports] = useState<ReportRow[]>([]);
  const [loading, setLoading] = useState(true);

  const [vmSaving, setVmSaving] = useState<Record<string, boolean>>({});

  // Search + filter
  const [q, setQ] = useState("");
  const [vmFilter, setVmFilter] = useState<VMFilter>("all");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("desc");

  // selection
  const [selected, setSelected] = useState<Record<string, boolean>>({});
  const [selMenuOpen, setSelMenuOpen] = useState(false);
  const selMenuRef = useRef<HTMLDivElement>(null);

  // watermark
  const [wmEnabled, setWmEnabled] = useState(true);
  const [wmText, setWmText] = useState("");
  const [wmDirty, setWmDirty] = useState(false);

  // Export modal
  const [exportModalOpen, setExportModalOpen] = useState(false);
  const [exportFormat, setExportFormat] = useState<ExportFormat>("docx");
  const [exportMode, setExportMode] = useState<ExportMode>("listed");
  const [exportName, setExportName] = useState("");
  const [stageRanges, setStageRanges] = useState("1-12,13-14,15-25");
  const [includePhotos, setIncludePhotos] = useState(true);

  // Download progress modal
  const [dlOpen, setDlOpen] = useState(false);
  const [dlTitle, setDlTitle] = useState("");
  const [dlError, setDlError] = useState<string | null>(null);
  const [dlDone, setDlDone] = useState(false);
  const [dlSecondsLeft, setDlSecondsLeft] = useState(0);
  const [preparedFiles, setPreparedFiles] = useState<PreparedFile[]>([]);

  const projectName = projectNameOf(project);

  useEffect(() => {
    const def = `CONFIDENTIAL REPORT: ${projectName}`;
    setWmText((prev) => (wmDirty ? prev : prev?.trim() ? prev : def));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projectName]);

  const watermarkOpts: WatermarkOpts = useMemo(() => {
    const def = `CONFIDENTIAL REPORT: ${projectName}`;
    return { enabled: wmEnabled, text: (wmText || def).trim() };
  }, [wmEnabled, wmText, projectName]);

  const fetchReports = async (searchText: string, dir: "asc" | "desc", filter: VMFilter) => {
    if (!projectId) return;

    // ✅ Select difficulty from DB
    let query = supabase
      .from("reports")
      .select("id, project_id, category, description, created_at, difficulty")
      .eq("project_id", projectId)
      .order("created_at", { ascending: dir === "asc" });

    // ✅ filter on difficulty
    if (filter === "unset") {
      query = query.is("difficulty", null);
    } else if (filter !== "all") {
      query = query.eq("difficulty", filter);
    }

    const sText = searchText.trim();
    if (sText) {
      query = query.textSearch("search_tsv", sText, {
        type: "websearch",
        config: "simple",
      });
    }

    const { data, error } = await query;
    if (error) throw error;
    setReports((data || []) as ReportRow[]);
  };

  const load = async () => {
    if (!projectId) return;
    setLoading(true);
    try {
      const { data: pData, error: pErr } = await supabase
        .from("projects")
        .select("*")
        .eq("id", projectId)
        .single();
      if (pErr) throw pErr;

      setProject(pData as ProjectRow);

      await fetchReports(q, sortDir, vmFilter);
      setSelected({});
    } catch (e: any) {
      alert(e?.message || String(e));
    } finally {
      setLoading(false);
    }
  };

  const onChangeVM = async (reportId: string, value: string) => {
    const next = vmDisplayToDb(value);
    setVmSaving((p) => ({ ...p, [reportId]: true }));

    // ✅ optimistic update difficulty
    setReports((prev) => prev.map((r) => (r.id === reportId ? { ...r, difficulty: next || null } : r)));

    try {
      await updateReportVM(reportId, next);
    } catch (e: any) {
      await fetchReports(q, sortDir, vmFilter);
      alert(e?.message || String(e));
    } finally {
      setVmSaving((p) => ({ ...p, [reportId]: false }));
    }
  };

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projectId]);

  useEffect(() => {
    if (!projectId) return;

    const t = setTimeout(() => {
      fetchReports(q, sortDir, vmFilter).catch((e: any) => alert(e?.message || String(e)));
      setSelected({});
    }, 250);

    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [q, sortDir, vmFilter, projectId]);

  // close selection menu on outside click
  useEffect(() => {
    function onDown(e: MouseEvent) {
      if (!selMenuRef.current) return;
      if (!selMenuRef.current.contains(e.target as Node)) setSelMenuOpen(false);
    }
    function onKey(e: KeyboardEvent) {
      if (e.key === "Escape") setSelMenuOpen(false);
    }
    if (selMenuOpen) {
      document.addEventListener("mousedown", onDown);
      document.addEventListener("keydown", onKey);
    }
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [selMenuOpen]);

  const filteredSortedReports = useMemo(() => reports, [reports]);

  const stats = useMemo(() => {
    const shown = filteredSortedReports.length;

    const last = filteredSortedReports.length
      ? new Date(
          [...filteredSortedReports].sort(
            (a, b) => new Date(b.created_at).getTime() - new Date(a.created_at).getTime()
          )[0].created_at
        ).toLocaleString()
      : "—";

    const selectedCount = filteredSortedReports.filter((r) => selected[r.id]).length;
    return { shown, last, selectedCount };
  }, [filteredSortedReports, selected]);

  const selectedIdsInOrder = useMemo(() => {
    return filteredSortedReports.filter((r) => selected[r.id]).map((r) => r.id);
  }, [filteredSortedReports, selected]);

  const toggleOne = (id: string) => setSelected((prev) => ({ ...prev, [id]: !prev[id] }));

  const selectAllVisible = () => {
    const next: Record<string, boolean> = { ...selected };
    for (const r of filteredSortedReports) next[r.id] = true;
    setSelected(next);
    setSelMenuOpen(false);
  };

  const clearSelection = () => {
    setSelected({});
    setSelMenuOpen(false);
  };

  // ========= Export modal helpers =========
  const openExportModal = () => {
    const baseListed = `${projectName}-${vmFilterLabel(vmFilter)}-${stats.shown}`;
    const baseSelectedOne = `${projectName}-SELECTED-${stats.selectedCount}`;
    const baseSelectedSplit = `${projectName}`;
    const baseAll = `${projectName}-ALL-REPORTS`;

    const defName =
      exportMode === "listed"
        ? baseListed
        : exportMode === "selectedOne"
          ? baseSelectedOne
          : exportMode === "selectedSplit"
            ? baseSelectedSplit
            : baseAll;

    setExportName((p) => (p?.trim() ? p : defName));
    setExportModalOpen(true);
  };

  // If user switches to GPX, force allowed modes
  useEffect(() => {
    if (!exportModalOpen) return;

    if (exportFormat === "gpx") {
      if (exportMode === "listed" || exportMode === "selectedSplit") {
        setExportMode(stats.selectedCount ? "selectedOne" : "all");
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [exportFormat, exportModalOpen]);

  const closeDl = () => {
    setDlOpen(false);
    setDlTitle("");
    setDlError(null);
    setDlDone(false);
    setDlSecondsLeft(0);
    setPreparedFiles([]);
  };

  // countdown timer (reverse)
  useEffect(() => {
    if (!dlOpen) return;
    if (dlDone) return;

    const t = setInterval(() => {
      setDlSecondsLeft((p) => (p > 0 ? p - 1 : 0));
    }, 1000);

    return () => clearInterval(t);
  }, [dlOpen, dlDone]);

  const startDlUI = (title: string, estSeconds: number) => {
    setDlTitle(title);
    setDlError(null);
    setDlDone(false);
    setPreparedFiles([]);
    setDlSecondsLeft(estSeconds);
    setDlOpen(true);
  };

  const runExport = async () => {
    if (!projectId) return;

    // validate selected
    if (exportMode === "selectedOne" && stats.selectedCount === 0) {
      alert("Please select at least 1 report.");
      return;
    }

    setExportModalOpen(false);

    try {
      // ✅ GPX branch (only all / selectedOne)
      if (exportFormat === "gpx") {
        const base = sanitizeFileBaseName(exportName || projectName);
        const fileName = `${base}.gpx`;

        const countForEst = exportMode === "selectedOne" ? stats.selectedCount : Math.max(stats.shown, 1);
        const est = estimateSeconds(exportMode, countForEst, false);
        startDlUI("Preparing GPX export…", est);

        if (exportMode === "all") {
          const { blob, fileName: fn } = await generateProjectGPX(supabase, projectId, {
            name: exportName || projectName,
            fileName,
          });
          setPreparedFiles([{ fileName: fn, blob }]);
          setDlDone(true);
          return;
        }

        if (exportMode === "selectedOne") {
          const ids = selectedIdsInOrder;
          const { blob, fileName: fn } = await generateProjectGPXByReportIds(
            supabase,
            projectId,
            ids,
            { name: exportName || projectName, fileName }
          );
          setPreparedFiles([{ fileName: fn, blob }]);
          setDlDone(true);
          return;
        }

        throw new Error("GPX supports only: Selected reports or All reports.");
      }

      // ✅ DOCX branch (unchanged)
      const count =
        exportMode === "listed"
          ? stats.shown
          : exportMode === "selectedOne" || exportMode === "selectedSplit"
            ? stats.selectedCount
            : Math.max(stats.shown, 1);

      const est = estimateSeconds(exportMode, count, includePhotos);
      startDlUI("Preparing DOCX export…", est);

      const wm = watermarkOpts.enabled ? watermarkOpts : { enabled: false, text: "" };

      if (exportMode === "all") {
        const fileName = `${sanitizeFileBaseName(exportName || `${projectName}-ALL-REPORTS`)}.docx`;
        const { blob } = await generateProjectDOCX(supabase, projectId, {
          includePhotos,
          fileName,
          watermark: wm as any,
        });
        setPreparedFiles([{ fileName, blob }]);
        setDlDone(true);
        return;
      }

      if (exportMode === "listed") {
        const ids = filteredSortedReports.map((r) => r.id);
        if (!ids.length) throw new Error("No reports available to export.");

        const fileName = `${sanitizeFileBaseName(
          exportName || `${projectName}-${vmFilterLabel(vmFilter)}-${ids.length}`
        )}.docx`;

        const { blob } = await generateProjectDOCXByReportIds(supabase, projectId, ids, {
          includePhotos,
          fileName,
          watermark: wm as any,
        });

        setPreparedFiles([{ fileName, blob }]);
        setDlDone(true);
        return;
      }

      if (exportMode === "selectedOne") {
        const ids = selectedIdsInOrder;
        const fileName = `${sanitizeFileBaseName(exportName || `${projectName}-SELECTED-${ids.length}`)}.docx`;

        const { blob } = await generateProjectDOCXByReportIds(supabase, projectId, ids, {
          includePhotos,
          fileName,
          watermark: wm as any,
        });

        setPreparedFiles([{ fileName, blob }]);
        setDlDone(true);
        return;
      }

      if (exportMode === "selectedSplit") {
        const ids = selectedIdsInOrder;
        const stages = parseStageRanges(stageRanges, ids.length);

        if (!stages.length) {
          throw new Error(`Invalid stage ranges.\nExample: "1-12,13-14,15-25"\nTotal selected: ${ids.length}`);
        }

        const base = sanitizeFileBaseName(exportName || projectName);
        const files: PreparedFile[] = [];

        for (const st of stages) {
          const subset = ids.slice(st.from - 1, st.to);
          if (!subset.length) continue;

          const fileName = `${base}-${st.label}.docx`;
          const { blob } = await generateProjectDOCXByReportIds(supabase, projectId, subset, {
            includePhotos,
            fileName,
            watermark: wm as any,
          });

          files.push({ fileName, blob });
        }

        if (!files.length) throw new Error("No stage files generated (check your stage ranges).");

        setPreparedFiles(files);
        setDlDone(true);
      }
    } catch (e: any) {
      setDlError(e?.message || String(e));
      setDlDone(true);
    }
  };

  return (
    <div style={styles.containerFluid}>
      <div style={styles.pageInner}>
        {/* ========= EXPORT MODAL ========= */}
        {exportModalOpen && (
          <div style={styles.modalOverlay} onMouseDown={() => setExportModalOpen(false)}>
            <div
              style={{ ...styles.modalCard, width: "min(720px, 96vw)" }}
              onMouseDown={(e) => e.stopPropagation()}
              role="dialog"
              aria-modal="true"
              aria-label="Export"
            >
              <div style={styles.modalTitle}>Export</div>

              <div style={styles.modalHint}>
                Generate file first (with an estimated timer), then click <b>Download now</b>.
              </div>

              <div style={{ marginTop: 4 }}>
                <div style={styles.routeLabel}>Format</div>
                <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                  <label style={styles.radioRow}>
                    <input
                      type="radio"
                      name="fmt"
                      checked={exportFormat === "docx"}
                      onChange={() => setExportFormat("docx")}
                    />
                    <div>
                      <div style={styles.radioTitle}>DOCX</div>
                      <div style={styles.radioSub}>Table export with photos + watermark</div>
                    </div>
                  </label>

                  <label style={styles.radioRow}>
                    <input
                      type="radio"
                      name="fmt"
                      checked={exportFormat === "gpx"}
                      onChange={() => setExportFormat("gpx")}
                    />
                    <div>
                      <div style={styles.radioTitle}>GPX</div>
                      <div style={styles.radioSub}>NE coordinate track points only</div>
                    </div>
                  </label>
                </div>
              </div>

              {/* Mode + options */}
              <div style={styles.exportGrid}>
                <div style={styles.routeLabel}>What to export</div>

                <div style={{ display: "grid", gap: 10 }}>
                  {exportFormat === "gpx" ? (
                    <>
                      <label
                        style={{
                          ...styles.radioRow,
                          opacity: stats.selectedCount ? 1 : 0.5,
                        }}
                      >
                        <input
                          type="radio"
                          name="mode"
                          disabled={!stats.selectedCount}
                          checked={exportMode === "selectedOne"}
                          onChange={() => setExportMode("selectedOne")}
                        />
                        <div>
                          <div style={styles.radioTitle}>Selected reports (one GPX)</div>
                          <div style={styles.radioSub}>Exports {stats.selectedCount} selected reports</div>
                        </div>
                      </label>

                      <label style={styles.radioRow}>
                        <input type="radio" name="mode" checked={exportMode === "all"} onChange={() => setExportMode("all")} />
                        <div>
                          <div style={styles.radioTitle}>All reports (Project)</div>
                          <div style={styles.radioSub}>Exports every report in this project</div>
                        </div>
                      </label>
                    </>
                  ) : (
                    <>
                      <label style={styles.radioRow}>
                        <input type="radio" name="mode" checked={exportMode === "listed"} onChange={() => setExportMode("listed")} />
                        <div>
                          <div style={styles.radioTitle}>Listed (current filter/search)</div>
                          <div style={styles.radioSub}>Exports {stats.shown} reports currently shown</div>
                        </div>
                      </label>

                      <label style={{ ...styles.radioRow, opacity: stats.selectedCount ? 1 : 0.5 }}>
                        <input
                          type="radio"
                          name="mode"
                          disabled={!stats.selectedCount}
                          checked={exportMode === "selectedOne"}
                          onChange={() => setExportMode("selectedOne")}
                        />
                        <div>
                          <div style={styles.radioTitle}>Selected (one DOCX)</div>
                          <div style={styles.radioSub}>Exports {stats.selectedCount} selected reports into one file</div>
                        </div>
                      </label>

                      <label style={{ ...styles.radioRow, opacity: stats.selectedCount ? 1 : 0.5 }}>
                        <input
                          type="radio"
                          name="mode"
                          disabled={!stats.selectedCount}
                          checked={exportMode === "selectedSplit"}
                          onChange={() => setExportMode("selectedSplit")}
                        />
                        <div>
                          <div style={styles.radioTitle}>Selected (split by stages)</div>
                          <div style={styles.radioSub}>Generates multiple DOCX files (A, B, C…)</div>
                        </div>
                      </label>

                      <label style={styles.radioRow}>
                        <input type="radio" name="mode" checked={exportMode === "all"} onChange={() => setExportMode("all")} />
                        <div>
                          <div style={styles.radioTitle}>All reports (Project)</div>
                          <div style={styles.radioSub}>Exports every report in this project</div>
                        </div>
                      </label>
                    </>
                  )}
                </div>

                {/* Name */}
                <div style={styles.routeLabel}>File name</div>
                <input style={styles.input} value={exportName} onChange={(e) => setExportName(e.target.value)} placeholder="Example: TSPL to Nallur" />

                {/* DOCX only: Split ranges */}
                <div style={styles.routeLabel}>Stage split</div>
                <input
                  style={{ ...styles.input, opacity: exportFormat === "docx" && exportMode === "selectedSplit" ? 1 : 0.5 }}
                  disabled={exportFormat !== "docx" || exportMode !== "selectedSplit"}
                  value={stageRanges}
                  onChange={(e) => setStageRanges(e.target.value)}
                  placeholder='Example: "1-12,13-14,15-25"'
                />

                {/* DOCX only: Options */}
                <div style={styles.routeLabel}>Options</div>
                {exportFormat === "docx" ? (
                  <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
                    <label style={{ display: "inline-flex", gap: 8, alignItems: "center", fontWeight: 800 }}>
                      <input type="checkbox" checked={includePhotos} onChange={(e) => setIncludePhotos(e.target.checked)} style={{ width: 16, height: 16 }} />
                      Include photos
                    </label>

                    <label style={{ display: "inline-flex", gap: 8, alignItems: "center", fontWeight: 800 }}>
                      <input type="checkbox" checked={wmEnabled} onChange={(e) => setWmEnabled(e.target.checked)} style={{ width: 16, height: 16 }} />
                      Watermark
                    </label>
                  </div>
                ) : (
                  <div style={{ fontSize: 12, fontWeight: 800, color: "#667085" }}>
                    GPX exports only NE coordinate points (no photos / watermark).
                  </div>
                )}

                {/* DOCX only: Watermark text */}
                <div style={styles.routeLabel}>Watermark text</div>
                <input
                  style={{ ...styles.input, opacity: exportFormat === "docx" && wmEnabled ? 1 : 0.5 }}
                  value={wmText}
                  disabled={exportFormat !== "docx" || !wmEnabled}
                  placeholder={`CONFIDENTIAL REPORT: ${projectName}`}
                  onChange={(e) => {
                    setWmDirty(true);
                    setWmText(e.target.value);
                  }}
                />
              </div>

              <div style={styles.modalActions}>
                <button style={styles.btnGhost} onClick={() => setExportModalOpen(false)}>
                  Cancel
                </button>
                <button style={styles.btnPrimary} onClick={runExport}>
                  Generate
                </button>
              </div>

              {exportFormat === "docx" ? (
                <div style={styles.modalNote}>
                  Tip: If export is slow, try disabling <b>Include photos</b>.
                </div>
              ) : (
                <div style={styles.modalNote}>Tip: Select only needed reports for smaller GPX.</div>
              )}
            </div>
          </div>
        )}

        {/* ========= DOWNLOAD PROGRESS MODAL ========= */}
        {dlOpen && (
          <div style={styles.modalOverlay} onMouseDown={() => (dlDone ? closeDl() : null)}>
            <div
              style={{ ...styles.modalCard, width: "min(640px, 96vw)" }}
              onMouseDown={(e) => e.stopPropagation()}
              role="dialog"
              aria-modal="true"
              aria-label="Download progress"
            >
              <div style={styles.modalTitle}>{dlTitle || "Working…"}</div>

              {!dlDone ? (
                <>
                  <div style={styles.modalHint}>
                    Estimated remaining: <b>{dlSecondsLeft}s</b>
                    {dlSecondsLeft === 0 ? <span style={{ marginLeft: 6, color: "#b42318" }}>(still working…)</span> : null}
                  </div>
                  <div style={styles.progressBarOuter}>
                    <div style={styles.progressBarInner} />
                  </div>
                  <div style={styles.modalNote}>Please keep this tab open until generation finishes.</div>
                </>
              ) : dlError ? (
                <>
                  <div style={{ ...styles.modalHint, color: "#b42318" }}>{dlError}</div>
                  <div style={styles.modalActions}>
                    <button style={styles.btnPrimary} onClick={closeDl}>
                      Close
                    </button>
                  </div>
                </>
              ) : (
                <>
                  <div style={styles.modalHint}>Ready. Click download below.</div>

                  <div style={{ display: "grid", gap: 10, marginTop: 8 }}>
                    {preparedFiles.map((f) => (
                      <button key={f.fileName} style={styles.btnPrimary} onClick={() => downloadBlob(f.blob, f.fileName)} title="Download now">
                        Download: {f.fileName}
                      </button>
                    ))}
                  </div>

                  <div style={styles.modalActions}>
                    <button style={styles.btnGhost} onClick={closeDl}>
                      Close
                    </button>
                  </div>

                  <div style={styles.modalNote}>
                    If your browser blocks multiple downloads (split stages), click each file button one-by-one.
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {/* ========= HEADER ========= */}
        <div style={styles.headerCard}>
          <div style={{ display: "grid", gap: 6, minWidth: 240 }}>
            <Link href="/projects" style={styles.backLink}>
              ← Back to Projects
            </Link>

            <div style={styles.title}>{projectName}</div>

            <div style={styles.metaRow}>
              <span style={styles.pill}>Project ID: {projectId}</span>
              <span style={styles.pill}>Showing: {stats.shown}</span>
              <span style={styles.pill}>Filter: {vmFilterLabel(vmFilter)}</span>
              <span style={styles.pill}>Last: {stats.last}</span>
              <span style={styles.pill}>Selected: {stats.selectedCount}</span>
            </div>
          </div>

          <div style={styles.actions}>
            <button style={styles.btnGhost} onClick={load} disabled={loading}>
              {loading ? "Refreshing..." : "Refresh"}
            </button>

            <button
              style={styles.btnGhost}
              onClick={() => setSortDir((p) => (p === "asc" ? "desc" : "asc"))}
              disabled={loading}
              title="Toggle ascending/descending"
            >
              Sort: {sortDir === "asc" ? "Ascending" : "Descending"}
            </button>

            <button style={styles.btnPrimary} onClick={openExportModal} disabled={loading}>
              Export
            </button>
          </div>
        </div>

        {/* ========= REPORTS CONTROLS ========= */}
        <div style={styles.card}>
          <div style={styles.cardHeader}>
            <div style={styles.cardTitle}>Reports</div>
            <div style={styles.cardHint}>{q.trim() ? `Results for “${q.trim()}”` : `Showing ${stats.shown} reports`}</div>
          </div>

          <div style={styles.controlsRow}>
            <div style={{ position: "relative", flex: 1, minWidth: 280 }}>
              <input
                style={{ ...styles.input, paddingRight: 38 }}
                placeholder='Search (FTS): try "bridge", "culvert", "tspl", "red"...'
                value={q}
                onChange={(e) => setQ(e.target.value)}
              />
              {q ? (
                <button style={styles.inputClearBtn} onClick={() => setQ("")} title="Clear search">
                  ×
                </button>
              ) : null}
            </div>

            <select value={vmFilter} onChange={(e) => setVmFilter(e.target.value as VMFilter)} style={styles.select} title="Filter by route difficulty">
              <option value="all">Difficulty: All</option>
              <option value="green">Green</option>
              <option value="yellow">Yellow</option>
              <option value="red">Red</option>
              <option value="unset">Not set</option>
            </select>

            <button
              style={styles.btnGhost}
              onClick={() => {
                setQ("");
                setVmFilter("all");
              }}
              disabled={!q && vmFilter === "all"}
              title="Clear search + filter"
            >
              Clear all
            </button>

            {/* Selection dropdown */}
            <div ref={selMenuRef} style={{ position: "relative" }}>
              <button style={styles.btnGhost} onClick={() => setSelMenuOpen((v) => !v)} title="Selection actions">
                Selection ▾
              </button>

              {selMenuOpen && (
                <div style={styles.menu}>
                  <button style={styles.menuItem} onClick={selectAllVisible} disabled={!stats.shown}>
                    Select all listed
                  </button>
                  <button style={styles.menuItem} onClick={clearSelection} disabled={!stats.selectedCount}>
                    Clear selection
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* ========= TABLE ========= */}
        {loading ? (
          <div style={styles.stateCard}>Loading...</div>
        ) : filteredSortedReports.length === 0 ? (
          <div style={styles.stateCard}>
            <div style={{ fontWeight: 900, color: "#101828" }}>No reports found</div>
            <div style={{ marginTop: 6, color: "#667085", fontWeight: 700 }}>Try a different search keyword or filter.</div>
          </div>
        ) : (
          <div style={styles.tableCard}>
            <div style={styles.tableWrapNoScroll}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th className="col-idx" style={styles.th}>#</th>
                    <th className="col-sel" style={styles.th}>Select</th>
                    <th className="col-cat" style={styles.th}>Category</th>
                    <th className="col-desc" style={styles.th}>Description</th>
                    <th className="col-created" style={styles.th}>Created</th>
                    <th className="col-id" style={styles.th}>Report ID</th>
                    <th className="col-vm" style={styles.th}>Route difficulty</th>
                    <th className="col-act" style={{ ...styles.th, textAlign: "right" }}>Actions</th>
                  </tr>
                </thead>

                <tbody>
                  {filteredSortedReports.map((r, i) => {
                    const created = r.created_at ? new Date(r.created_at).toLocaleString() : "—";
                    const desc = displayDescription((r.description || "").trim());
                    const shortId = r.id ? `${r.id.slice(0, 8)}...` : "—";

                    // ✅ read difficulty from DB
                    const vmValue = normalizeVM(r.difficulty);

                    return (
                      <tr key={r.id}>
                        <td className="col-idx" style={styles.td}>{i + 1}</td>

                        <td className="col-sel" style={styles.td}>
                          <input
                            type="checkbox"
                            checked={!!selected[r.id]}
                            onChange={() => toggleOne(r.id)}
                            style={{ width: 18, height: 18, cursor: "pointer" }}
                          />
                        </td>

                        <td className="col-cat" style={styles.td}>
                          <div style={styles.catTitle}>{r.category || "Report"}</div>
                          <div style={styles.subtle}>Includes photos</div>
                        </td>

                        <td className="col-desc" style={styles.td}>
                          <div style={styles.descCell}>
                            {desc ? desc : <span style={{ color: "#98A2B3", fontWeight: 800 }}>No description</span>}
                          </div>
                        </td>

                        <td className="col-created" style={styles.td}>
                          <span style={styles.mutedWrap}>{created}</span>
                        </td>

                        <td className="col-id" style={styles.td}>
                          <span style={styles.codePillWrap} title={r.id}>{shortId}</span>
                        </td>

                        <td className="col-vm" style={styles.td}>
                          <select
                            value={vmValue || ""}
                            disabled={!!vmSaving[r.id]}
                            onChange={(e) => onChangeVM(r.id, e.target.value)}
                            style={{
                              height: 38,
                              borderRadius: 14,
                              border: "1px solid #EAECF0",
                              padding: "0 12px",
                              fontWeight: 950,
                              background:
                                vmValue === "green" ? "#EAFBF0" : vmValue === "yellow" ? "#FFFBEB" : vmValue === "red" ? "#FEF2F2" : "#F2F4F7",
                              color:
                                vmValue === "green" ? "#067647" : vmValue === "yellow" ? "#92400E" : vmValue === "red" ? "#B42318" : "#475467",
                              cursor: vmSaving[r.id] ? "not-allowed" : "pointer",
                              outline: "none",
                              minWidth: 130,
                              opacity: vmSaving[r.id] ? 0.7 : 1,
                            }}
                            title="Update route difficulty"
                          >
                            <option value="">Not set</option>
                            <option value="green">Green</option>
                            <option value="yellow">Yellow</option>
                            <option value="red">Red</option>
                          </select>
                        </td>

                        <td className="col-act" style={{ ...styles.td, textAlign: "right" }}>
                          <Link href={`/reports/${r.id}`} style={styles.btnOpen} title="Open report">
                            Open
                          </Link>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <style jsx>{`
              table { width: 100%; }
              .col-desc,.col-created,.col-id,.col-vm { word-break: break-word; }
              .col-idx { width: 44px; }
              .col-sel { width: 72px; }
              .col-cat { width: 180px; }
              .col-created { width: 190px; }
              .col-id { width: 120px; }
              .col-vm { width: 150px; }
              .col-act { width: 110px; }
              @media (max-width: 1200px) { .col-desc { display: none; } }
              @media (max-width: 992px) { .col-id { display: none; } }
              @media (max-width: 820px) { .col-created { display: none; } }
              @media (max-width: 640px) {
                .col-idx { display: none; }
                .col-sel { width: 64px; }
                .col-cat { width: 150px; }
                .col-vm { width: auto; }
                .col-act { width: 90px; }
              }
            `}</style>
          </div>
        )}
      </div>
    </div>
  );
}

// ✅ keep your existing styles object EXACTLY as-is below (unchanged)
const styles: Record<string, React.CSSProperties> = {
  containerFluid: { background: "#F7F8FA", minHeight: "100vh", padding: "18px 18px" },
  pageInner: {
    width: "100%",
    margin: 0,
    display: "grid",
    gap: 14,
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
  },

  headerCard: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    padding: 16,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: 14,
    flexWrap: "wrap",
  },
  backLink: { textDecoration: "none", color: "#344054", fontWeight: 800, fontSize: 13 },
  title: { fontSize: 22, fontWeight: 900, color: "#101828", lineHeight: 1.2 },
  metaRow: { display: "flex", gap: 8, flexWrap: "wrap", marginTop: 2 },
  pill: {
    fontSize: 12,
    fontWeight: 800,
    color: "#475467",
    background: "#F2F4F7",
    border: "1px solid #EAECF0",
    borderRadius: 999,
    padding: "6px 10px",
  },

  actions: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },
  btnPrimary: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #111",
    background: "#111",
    color: "#fff",
    cursor: "pointer",
    fontWeight: 900,
    fontSize: 13,
    height: 40,
    whiteSpace: "nowrap",
  },
  btnGhost: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 900,
    fontSize: 13,
    color: "#344054",
    height: 40,
    whiteSpace: "nowrap",
  },

  select: {
    height: 40,
    borderRadius: 12,
    border: "1px solid #EAECF0",
    padding: "0 10px",
    fontWeight: 900,
    color: "#101828",
    background: "#fff",
    minWidth: 180,
  },

  card: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    padding: 14,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },
  cardHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "baseline",
    gap: 10,
    marginBottom: 10,
  },
  cardTitle: { fontSize: 14, fontWeight: 900, color: "#101828" },
  cardHint: { fontSize: 12, fontWeight: 800, color: "#667085" },

  controlsRow: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" },
  input: {
    flex: 1,
    minWidth: 220,
    padding: "12px 14px",
    borderRadius: 12,
    border: "1px solid #EAECF0",
    outline: "none",
    fontSize: 14,
    background: "#fff",
  },
  inputClearBtn: {
    position: "absolute",
    right: 8,
    top: "50%",
    transform: "translateY(-50%)",
    width: 28,
    height: 28,
    borderRadius: 10,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 900,
    color: "#667085",
    lineHeight: "26px",
  },

  stateCard: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    padding: 18,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },

  tableCard: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 18,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
    overflow: "hidden",
  },
  tableWrapNoScroll: { width: "100%", overflowX: "hidden" },
  table: { width: "100%", borderCollapse: "separate", borderSpacing: 0, tableLayout: "fixed" },
  th: {
    textAlign: "left",
    fontSize: 12,
    letterSpacing: 0.2,
    fontWeight: 900,
    color: "#475467",
    background: "#F9FAFB",
    borderBottom: "1px solid #EAECF0",
    padding: "12px 12px",
    position: "sticky",
    top: 0,
    zIndex: 1,
    whiteSpace: "nowrap",
  },
  td: { padding: "12px 12px", borderBottom: "1px solid #F2F4F7", verticalAlign: "top", fontSize: 13, color: "#101828", overflow: "hidden" },
  catTitle: { fontWeight: 900, fontSize: 13, color: "#101828" },
  descCell: { color: "#475467", lineHeight: 1.45, fontWeight: 700, whiteSpace: "normal", wordBreak: "break-word" },
  subtle: { fontSize: 12, fontWeight: 800, color: "#667085" },
  mutedWrap: { fontSize: 12, fontWeight: 800, color: "#667085", whiteSpace: "normal", wordBreak: "break-word" },
  codePillWrap: {
    display: "inline-block",
    fontSize: 12,
    fontWeight: 900,
    color: "#344054",
    background: "#F2F4F7",
    border: "1px solid #EAECF0",
    borderRadius: 999,
    padding: "6px 10px",
    whiteSpace: "nowrap",
  },
  btnOpen: {
    textDecoration: "none",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "9px 12px",
    borderRadius: 12,
    border: "1px solid #B2DDFF",
    background: "#EFF8FF",
    color: "#175CD3",
    fontWeight: 900,
    fontSize: 12,
    height: 36,
    whiteSpace: "nowrap",
  },

  // menus
  menu: {
    position: "absolute",
    top: "calc(100% + 8px)",
    right: 0,
    width: 220,
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 14,
    boxShadow: "0 12px 32px rgba(16,24,40,0.12)",
    padding: 6,
    zIndex: 50,
  },
  menuItem: {
    width: "100%",
    textAlign: "left",
    padding: "10px 10px",
    borderRadius: 12,
    border: "none",
    background: "transparent",
    cursor: "pointer",
    fontWeight: 900,
    fontSize: 13,
    color: "#101828",
  },

  // modal
  modalOverlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(16,24,40,0.45)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 14,
    zIndex: 9999,
  },
  modalCard: {
    width: "min(520px, 96vw)",
    background: "#fff",
    borderRadius: 18,
    border: "1px solid #EAECF0",
    boxShadow: "0 20px 60px rgba(16,24,40,0.25)",
    padding: 16,
    display: "grid",
    gap: 10,
  },
  modalTitle: { fontSize: 16, fontWeight: 950, color: "#101828" },
  modalHint: { fontSize: 12, fontWeight: 800, color: "#667085", lineHeight: 1.35 },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 6 },
  modalNote: { fontSize: 12, fontWeight: 750, color: "#667085" },

  exportGrid: { display: "grid", gridTemplateColumns: "160px 1fr", gap: 10, alignItems: "start", marginTop: 6 },
  routeLabel: { fontSize: 12, fontWeight: 900, color: "#475467" },

  radioRow: {
    display: "flex",
    gap: 10,
    alignItems: "flex-start",
    padding: "10px 10px",
    borderRadius: 14,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
  },
  radioTitle: { fontWeight: 950, color: "#101828", fontSize: 13 },
  radioSub: { fontWeight: 800, color: "#667085", fontSize: 12, marginTop: 2 },

  progressBarOuter: {
    height: 10,
    borderRadius: 999,
    background: "#F2F4F7",
    border: "1px solid #EAECF0",
    overflow: "hidden",
    marginTop: 6,
  },
  progressBarInner: {
    height: "100%",
    width: "65%",
    borderRadius: 999,
    background: "#111",
    animation: "pulse 1.1s ease-in-out infinite",
  },
};
