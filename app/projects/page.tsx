"use client";

import React, { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { supabase } from "../../lib/supabaseClient";
import { downloadProjectsCSV } from "../../lib/download";

type ProjectRow = {
  id: string;
  name?: string | null;
  title?: string | null;
  project_name?: string | null;
  created_at?: string | null;

  // ✅ optional if you have this column
  updated_at?: string | null;
  last_modified_at?: string | null;
  last_modified_by?: string | null; // user uuid
};

type ProfileRow = {
  id: string;
  full_name?: string | null;
  name?: string | null;
  email?: string | null;
};

export default function ProjectsPage() {
  const [projects, setProjects] = useState<ProjectRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [q, setQ] = useState("");
  const [exporting, setExporting] = useState(false);

  // ✅ last modified display (projectId -> name)
  const [lastModifiedMap, setLastModifiedMap] = useState<Record<string, string>>(
    {}
  );

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    if (!s) return projects;
    return projects.filter((p) => JSON.stringify(p).toLowerCase().includes(s));
  }, [projects, q]);

  const safeName = (p: ProjectRow) =>
    p.name || p.title || p.project_name || "Untitled Project";

  const load = async () => {
    setLoading(true);
    try {
      const { data, error } = await supabase
        .from("projects")
        .select("*")
        .order("created_at", { ascending: false });

      if (error) throw error;

      const rows = (data || []) as ProjectRow[];
      setProjects(rows);

      // ✅ Build last-modified names for all visible projects
      await hydrateLastModifiedNames(rows);
    } catch (e: any) {
      alert(e?.message || String(e));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
  }, []);

  const logout = async () => {
    await supabase.auth.signOut();
  };

  const exportCSV = async () => {
    if (!filtered.length) return;
    try {
      setExporting(true);
      await downloadProjectsCSV(filtered);
    } finally {
      setExporting(false);
    }
  };

  /**
   * ✅ LAST MODIFIED NAME RESOLUTION (robust)
   * Priority:
   * A) projects.last_modified_by -> profiles
   * B) latest row from project_modifications -> profiles
   * fallback: —
   */
 const hydrateLastModifiedNames = async (rows: ProjectRow[]) => {
  try {
    const userIds = Array.from(
      new Set(rows.map((r) => r.last_modified_by).filter(Boolean) as string[])
    );

    if (!userIds.length) {
      setLastModifiedMap({});
      return;
    }

    const { data: profiles, error } = await supabase
      .from("profiles")
      .select("id, full_name, name, email")
      .in("id", userIds);

    if (error) throw error;

    const userIdToName: Record<string, string> = {};
    (profiles || []).forEach((u: any) => {
      userIdToName[String(u.id)] =
        u?.full_name || u?.name || u?.email || String(u.id).slice(0, 8);
    });

    const map: Record<string, string> = {};
    rows.forEach((p) => {
      if (p.last_modified_by) map[p.id] = userIdToName[p.last_modified_by] || "—";
    });

    setLastModifiedMap(map);
  } catch {
    setLastModifiedMap({});
  }
};


  return (
    <div style={styles.page}>
      {/* Header */}
      <div style={styles.header}>
        <div>
          <div style={styles.title}>Projects</div>
          <div style={styles.subtitle}>
            Total: <b>{projects.length}</b> • Showing: <b>{filtered.length}</b>
          </div>
        </div>

        <div style={styles.headerRight}>
          <button
            style={{ ...styles.btnGhost, opacity: exporting ? 0.7 : 1 }}
            onClick={load}
            disabled={loading || exporting}
          >
            {loading ? "Refreshing..." : "Refresh"}
          </button>

          {/* Export group */}
          <div style={styles.exportGroup}>
            <button
              style={{ ...styles.btnPrimary, opacity: exporting ? 0.7 : 1 }}
              onClick={exportCSV}
              disabled={!filtered.length || exporting}
              title={!filtered.length ? "No projects to export" : "Export CSV"}
            >
              {exporting ? "Exporting..." : "Export CSV"}
            </button>
          </div>

          <button style={styles.btnDanger} onClick={logout}>
            Logout
          </button>
        </div>
      </div>

      {/* Search */}
      <div style={styles.searchBar}>
        <div style={styles.searchWrap}>
          <span style={styles.searchIcon}>⌕</span>
          <input
            style={styles.searchInput}
            placeholder="Search projects by name, id, any field..."
            value={q}
            onChange={(e) => setQ(e.target.value)}
          />
          {q ? (
            <button style={styles.clearBtn} onClick={() => setQ("")}>
              Clear
            </button>
          ) : null}
        </div>
      </div>

      {/* Content */}
      {loading ? (
        <div style={styles.stateBox}>Loading projects...</div>
      ) : filtered.length === 0 ? (
        <div style={styles.stateBox}>
          <div style={{ fontWeight: 700, marginBottom: 6 }}>No projects found</div>
          <div style={{ color: "#667085" }}>Try a different keyword or click Refresh.</div>
        </div>
      ) : (
        <div style={styles.grid}>
          {filtered.map((p) => {
            const name = safeName(p);
            const dt = p.created_at ? new Date(p.created_at).toLocaleString() : "";
            const modifiedBy = lastModifiedMap[p.id] || "—";

            return (
              <Link key={p.id} href={`/projects/${p.id}`} style={styles.card}>
                <div style={styles.cardTop}>
                  <div style={styles.cardTitle}>{name}</div>
                  <span style={styles.badge}>Open</span>
                </div>

                <div style={styles.metaRow}>
                  <span style={styles.metaLabel}>Project ID</span>
                  <span style={styles.metaValue} title={p.id}>
                    {p.id}
                  </span>
                </div>

                <div style={styles.metaRow}>
                  <span style={styles.metaLabel}>Created</span>
                  <span style={styles.metaValue}>{dt || "—"}</span>
                </div>

                {/* ✅ NEW: Last Modified By */}
                <div style={styles.metaRow}>
                  <span style={styles.metaLabel}>Last modified by</span>
                  <span style={styles.metaValue} title={modifiedBy}>
                    {modifiedBy}
                  </span>
                </div>

                <div style={styles.cardHint}>Click to view reports →</div>
              </Link>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ---------------- styles ---------------- */

const styles: Record<string, React.CSSProperties> = {
  page: {
    padding: 24,
    fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
    background: "#F7F8FA",
    minHeight: "100vh",
    maxWidth: 1400,
    margin: "0 auto",
  },

  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: 12,
    flexWrap: "wrap",
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 16,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },

  title: { fontSize: 22, fontWeight: 800, color: "#101828", lineHeight: 1.2 },
  subtitle: { fontSize: 13, color: "#667085", marginTop: 6 },

  headerRight: {
    display: "flex",
    gap: 10,
    alignItems: "center",
    flexWrap: "wrap",
    justifyContent: "flex-end",
  },

  exportGroup: {
    display: "flex",
    gap: 8,
    flexWrap: "wrap",
    alignItems: "center",
  },

  btnPrimary: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #111",
    background: "#111",
    color: "#fff",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 13,
  },

  btnGhost: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 13,
    color: "#344054",
  },

  btnDanger: {
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #FDA29B",
    background: "#FEF3F2",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 13,
    color: "#B42318",
  },

  searchBar: { marginTop: 14, marginBottom: 14 },

  searchWrap: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 14,
    padding: "10px 12px",
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },

  searchIcon: { color: "#667085", fontSize: 14 },

  searchInput: {
    flex: 1,
    border: "none",
    outline: "none",
    fontSize: 14,
    color: "#101828",
  },

  clearBtn: {
    padding: "8px 10px",
    borderRadius: 10,
    border: "1px solid #EAECF0",
    background: "#fff",
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 12,
    color: "#344054",
  },

  stateBox: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 18,
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },

  grid: {
    display: "grid",
    gridTemplateColumns: "repeat(4, minmax(0, 1fr))",
    gap: 12,
  },

  card: {
    background: "#fff",
    border: "1px solid #EAECF0",
    borderRadius: 16,
    padding: 14,
    textDecoration: "none",
    color: "#101828",
    boxShadow: "0 1px 2px rgba(16,24,40,0.06)",
  },

  cardTop: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: 10,
    marginBottom: 10,
  },

  cardTitle: {
    fontSize: 16,
    fontWeight: 800,
    lineHeight: 1.2,
  },

  badge: {
    fontSize: 12,
    fontWeight: 800,
    padding: "4px 10px",
    borderRadius: 999,
    border: "1px solid #D0D5DD",
    background: "#F9FAFB",
    color: "#344054",
    whiteSpace: "nowrap",
  },

  metaRow: {
    display: "flex",
    justifyContent: "space-between",
    gap: 10,
    padding: "8px 0",
    borderTop: "1px dashed #EAECF0",
  },

  metaLabel: { fontSize: 12, color: "#667085", fontWeight: 700 },
  metaValue: {
    fontSize: 12,
    color: "#101828",
    fontWeight: 700,
    maxWidth: 190,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },

  cardHint: {
    marginTop: 10,
    fontSize: 12,
    color: "#475467",
    fontWeight: 700,
  },
};
