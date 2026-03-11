"use client";

import React, { useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import * as XLSX from "xlsx";
import { supabase } from "../../lib/supabaseClient";
import { downloadProjectsCSV } from "../../lib/download";

type ProjectRow = {
  id: string;
  user_id?: string | null;
  name?: string | null;
  title?: string | null;
  project_name?: string | null;
  description?: string | null;
  created_at?: string | null;
  updated_at?: string | null;
  last_modified_by?: string | null;
  created_by?: string | null;
};

type ParsedPointRow = {
  point_key: string;
  latitude: number;
  longitude: number;
  category: string;
  normalizedCategory?: string | null;
  description?: string | null;
};

type ParsedImageMapRow = {
  file_name: string;
  point_key: string | null;
  image_key?: string | null;
};

type ImportSummary = {
  pointsRead: number;
  reportsCreatedOrUsed: number;
  imagesSelected: number;
  imagesUploaded: number;
  photosInserted: number;
  noGpsImages: string[];
  missingFilesInUpload: string[];
  extraFilesNotInMap: string[];
  errors: string[];
  duplicateSelectedImages: string[];
  duplicateMappingFiles: string[];
  missingPointKeysInPointsCsv: string[];
  invalidCategories: string[];
};

const CATEGORY_OPTIONS = [
  "Footpath Bridge",
  "Low Tension Cable",
  "High Tension Cable",
  "Towerline Cable",
  "Take Diversion",
  "Towerline",
  "Underpass Bridge",
  "Tree Branches",
  "Bridge",
  "Petrol bunk",
  "Signboard",
  "Electric Sign Board",
  "Camera Pole",
  "Toll Plaza",
  "Junction left",
  "Bend",
  "Junction right",
] as const;

const CATEGORY_ALIAS_MAP: Record<string, string> = {
  "footpath bridge": "Footpath Bridge",
  "footpathbridge": "Footpath Bridge",
  "foot bridge": "Footpath Bridge",
  "pedestrian bridge": "Footpath Bridge",

  "low tension cable": "Low Tension Cable",
  "low tension cables": "Low Tension Cable",
  "lt cable": "Low Tension Cable",
  "lt cables": "Low Tension Cable",
  "lt line": "Low Tension Cable",
  "lt lines": "Low Tension Cable",

  "high tension cable": "High Tension Cable",
  "high tension cables": "High Tension Cable",
  "ht cable": "High Tension Cable",
  "ht cables": "High Tension Cable",
  "ht line": "High Tension Cable",
  "ht lines": "High Tension Cable",

  "towerline cable": "Towerline Cable",
  "towerline cables": "Towerline Cable",
  "tower line cable": "Towerline Cable",
  "tower line cables": "Towerline Cable",

  "take diversion": "Take Diversion",
  diversion: "Take Diversion",
  diversions: "Take Diversion",
  "take diversions": "Take Diversion",

  towerline: "Towerline",
  "tower line": "Towerline",
  "tower lines": "Towerline",
  "transmission tower": "Towerline",

  "underpass bridge": "Underpass Bridge",
  underpass: "Underpass Bridge",
  "under bridge": "Underpass Bridge",

  "tree branches": "Tree Branches",
  "tree branch": "Tree Branches",
  branches: "Tree Branches",
  tree: "Tree Branches",

  bridge: "Bridge",
  bridges: "Bridge",
  flyover: "Bridge",

  "petrol bunk": "Petrol bunk",
  "petrol bunks": "Petrol bunk",
  "petrol pump": "Petrol bunk",
  "fuel station": "Petrol bunk",
  "fuel bunk": "Petrol bunk",

  signboard: "Signboard",
  signboards: "Signboard",
  "sign board": "Signboard",
  "sign boards": "Signboard",
  "road sign": "Signboard",

  "electric sign board": "Electric Sign Board",
  "electric signboard": "Electric Sign Board",
  "electric sign boards": "Electric Sign Board",
  "electrical sign board": "Electric Sign Board",
  "illuminated sign board": "Electric Sign Board",

  "camera pole": "Camera Pole",
  "camera poles": "Camera Pole",
  "cctv pole": "Camera Pole",
  "surveillance pole": "Camera Pole",

  "toll plaza": "Toll Plaza",
  "toll plazas": "Toll Plaza",
  toll: "Toll Plaza",

  "junction left": "Junction left",
  "left junction": "Junction left",
  "left turn junction": "Junction left",
  "left turn": "Junction left",

  bend: "Bend",
  bends: "Bend",
  curve: "Bend",
  curves: "Bend",

  "junction right": "Junction right",
  "right junction": "Junction right",
  "right turn junction": "Junction right",
  "right turn": "Junction right",
};

function normalizeHeader(h: string) {
  return h.trim().toLowerCase().replace(/\s+/g, "_");
}

function detectDelimiter(line: string) {
  return line.includes("\t") ? "\t" : ",";
}

function stripBom(text: string) {
  return text.replace(/^\uFEFF/, "");
}

function normalizeCategoryName(input: string) {
  const raw = String(input || "").trim();
  if (!raw) return "Unknown";

  const lowered = raw
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const exact = CATEGORY_OPTIONS.find((x) => x.toLowerCase() === lowered);
  if (exact) return exact;

  return CATEGORY_ALIAS_MAP[lowered] || raw;
}

function parseCSVLike(text: string) {
  const raw = stripBom(text).replace(/\r/g, "").trim();
  if (!raw) return { headers: [] as string[], rows: [] as string[][] };

  const lines = raw
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);

  if (!lines.length) return { headers: [] as string[], rows: [] as string[][] };

  const delim = detectDelimiter(lines[0]);
  const headers = lines[0].split(delim).map((h) => normalizeHeader(h));
  const rows = lines.slice(1).map((ln) => ln.split(delim).map((c) => c.trim()));

  return { headers, rows };
}

function colIndex(headers: string[], names: string[]) {
  for (const n of names) {
    const idx = headers.indexOf(n);
    if (idx >= 0) return idx;
  }
  return -1;
}

function isExcelFile(file: File) {
  const name = file.name.toLowerCase();
  return name.endsWith(".xlsx") || name.endsWith(".xls");
}

async function readWorkbookRows(file: File): Promise<string[][]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) return [];

  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
  }) as any[][];

  return rows.map((row) => row.map((cell) => String(cell ?? "").trim()));
}

async function readStructuredFile(file: File) {
  if (isExcelFile(file)) {
    const rows = await readWorkbookRows(file);
    if (!rows.length) return { headers: [] as string[], rows: [] as string[][] };

    const headers = rows[0].map((h) => normalizeHeader(String(h || "")));
    const dataRows = rows
      .slice(1)
      .filter((r) => r.some((c) => String(c || "").trim() !== ""))
      .map((r) => r.map((c) => String(c || "").trim()));

    return { headers, rows: dataRows };
  }

  const text = await readTextFile(file);
  return parseCSVLike(text);
}

async function parsePoints(file: File): Promise<ParsedPointRow[]> {
  const { headers, rows } = await readStructuredFile(file);
  if (!headers.length) return [];

  const iKey = colIndex(headers, ["point_key", "point", "seq", "key"]);
  const iLat = colIndex(headers, ["latitude", "lat"]);
  const iLon = colIndex(headers, ["longitude", "lon", "lng"]);
  const iCat = colIndex(headers, ["category", "report_category"]);
  const iDesc = colIndex(headers, ["description", "report_description", "desc"]);

  if (iKey < 0 || iLat < 0 || iLon < 0 || iCat < 0) {
    throw new Error(
      "Points file must include headers: point_key, latitude, longitude, category"
    );
  }

  const out: ParsedPointRow[] = [];

  for (const r of rows) {
    const key = (r[iKey] ?? "").trim();
    if (!key) continue;

    const lat = Number((r[iLat] ?? "").trim());
    const lon = Number((r[iLon] ?? "").trim());
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) continue;

    const rawCategory = (r[iCat] ?? "").trim();
    const finalCategory = normalizeCategoryName(rawCategory);

    out.push({
      point_key: key,
      latitude: lat,
      longitude: lon,
      category: finalCategory || "Unknown",
      normalizedCategory: finalCategory || null,
      description: iDesc >= 0 ? ((r[iDesc] ?? "").trim() || null) : null,
    });
  }

  return out;
}

async function parseImageMap(file: File): Promise<ParsedImageMapRow[]> {
  const { headers, rows } = await readStructuredFile(file);
  if (!headers.length) return [];

  const iFile = colIndex(headers, ["file_name", "filename", "file", "name"]);
  const iPoint = colIndex(headers, ["point_key", "point", "seq", "key"]);
  const iImgKey = colIndex(headers, ["image_key", "imagekey", "img_key", "imgkey"]);

  if (iFile < 0) {
    throw new Error("Images mapping file must include header: file_name");
  }

  const out: ParsedImageMapRow[] = [];
  for (const r of rows) {
    const file_name = (r[iFile] ?? "").trim();
    if (!file_name) continue;

    let point_key = iPoint >= 0 ? ((r[iPoint] ?? "").trim() || null) : null;
    if (point_key) {
      const up = point_key.toUpperCase();
      if (up === "NO_GPS" || up === "NOGPS" || up === "NULL" || up === "NONE") {
        point_key = null;
      }
    }

    out.push({
      file_name,
      point_key,
      image_key: iImgKey >= 0 ? ((r[iImgKey] ?? "").trim() || null) : null,
    });
  }

  return out;
}

async function readTextFile(file: File): Promise<string> {
  if (!file) throw new Error("No file selected.");

  try {
    return await file.text();
  } catch {
    return await new Promise((resolve, reject) => {
      try {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ""));
        reader.onerror = () =>
          reject(
            new Error(
              `Failed to read file "${file.name}". Please re-select the file from a normal local folder and try again.`
            )
          );
        reader.readAsText(file);
      } catch (err: any) {
        reject(
          new Error(
            `Failed to read file "${file.name}". ${err?.message || String(err)}`
          )
        );
      }
    });
  }
}

async function getImageSize(
  file: File
): Promise<{ width: number | null; height: number | null }> {
  try {
    const bmp = await createImageBitmap(file);
    const width = bmp.width ?? null;
    const height = bmp.height ?? null;
    bmp.close?.();
    return { width, height };
  } catch {
    return { width: null, height: null };
  }
}

async function uploadToBucket(bucket: string, path: string, file: File) {
  const { data, error } = await supabase.storage.from(bucket).upload(path, file, {
    cacheControl: "3600",
    upsert: true,
    contentType: file.type || undefined,
  });
  if (error) throw error;

  const { data: pub } = supabase.storage.from(bucket).getPublicUrl(data.path);
  return { path: data.path, publicUrl: pub.publicUrl };
}

async function mapLimit<T>(
  items: T[],
  limit: number,
  worker: (item: T, index: number) => Promise<void>
) {
  let nextIndex = 0;

  async function runner() {
    while (true) {
      const i = nextIndex++;
      if (i >= items.length) return;
      await worker(items[i], i);
    }
  }

  await Promise.all(
    Array.from({ length: Math.max(1, limit) }, () => runner())
  );
}

function getDuplicates(values: string[]) {
  const seen = new Set<string>();
  const dup = new Set<string>();

  for (const v of values) {
    const key = v.trim();
    if (!key) continue;
    if (seen.has(key)) dup.add(key);
    else seen.add(key);
  }

  return Array.from(dup).sort((a, b) => a.localeCompare(b));
}

export default function ProjectsPage() {
  const router = useRouter();

  const [projects, setProjects] = useState<ProjectRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [q, setQ] = useState("");
  const [exporting, setExporting] = useState(false);

  const [lastModifiedMap, setLastModifiedMap] = useState<Record<string, string>>({});

  const [newOpen, setNewOpen] = useState(false);
  const [newName, setNewName] = useState("");
  const [newDesc, setNewDesc] = useState("");
  const [creating, setCreating] = useState(false);

  const [bulkOpen, setBulkOpen] = useState(false);
  const [bulkProjectId, setBulkProjectId] = useState<string>("");
  const [bucketName, setBucketName] = useState<string>("reports");
  const [importing, setImporting] = useState(false);

  const [pointsFile, setPointsFile] = useState<File | null>(null);
  const [imageMapFile, setImageMapFile] = useState<File | null>(null);
  const [imageFiles, setImageFiles] = useState<File[]>([]);

  const pointsInputRef = useRef<HTMLInputElement | null>(null);
  const mapInputRef = useRef<HTMLInputElement | null>(null);
  const imagesInputRef = useRef<HTMLInputElement | null>(null);

  const [summary, setSummary] = useState<ImportSummary | null>(null);
  const [pointsPreview, setPointsPreview] = useState<string>("");
  const [imageMapPreview, setImageMapPreview] = useState<string>("");

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    if (!s) return projects;
    return projects.filter((p) => JSON.stringify(p).toLowerCase().includes(s));
  }, [projects, q]);

  const safeName = (p: ProjectRow) =>
    p.name || p.title || p.project_name || "Untitled Project";

  const redirectToLogin = () => {
    router.replace("/login");
  };

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
      await hydrateLastModifiedNames(rows);

      if (!bulkProjectId && rows.length) setBulkProjectId(rows[0].id);
    } catch (e: any) {
      const msg = String(e?.message || e || "").toLowerCase();
      if (msg.includes("auth") || msg.includes("session") || msg.includes("jwt")) {
        redirectToLogin();
        return;
      }
      alert(e?.message || String(e));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    const init = async () => {
      try {
        const {
          data: { session },
          error,
        } = await supabase.auth.getSession();

        if (error) throw error;

        if (!session?.user) {
          redirectToLogin();
          return;
        }

        await load();
      } catch {
        redirectToLogin();
      }
    };

    init();
  }, []);

  useEffect(() => {
    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event) => {
      if (event === "SIGNED_OUT") {
        redirectToLogin();
      }
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  const logout = async () => {
    await supabase.auth.signOut();
    redirectToLogin();
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

  const createProject = async () => {
    const name = newName.trim();
    if (!name) return alert("Project name is required");

    try {
      setCreating(true);

      const {
        data: { session },
        error: sessionErr,
      } = await supabase.auth.getSession();

      if (sessionErr) throw sessionErr;

      if (!session?.user) {
        redirectToLogin();
        return;
      }

      const user = session.user;

      const { error } = await supabase.from("projects").insert([
        {
          user_id: user.id,
          name,
          description: newDesc.trim() || null,
          created_by: user.id,
          last_modified_by: user.id,
        },
      ]);

      if (error) throw error;

      setNewOpen(false);
      setNewName("");
      setNewDesc("");
      await load();
    } catch (e: any) {
      const msg = String(e?.message || e || "").toLowerCase();
      if (msg.includes("auth") || msg.includes("session") || msg.includes("jwt")) {
        redirectToLogin();
        return;
      }
      alert(e?.message || String(e));
    } finally {
      setCreating(false);
    }
  };

  const resetBulk = () => {
    setPointsFile(null);
    setImageMapFile(null);
    setImageFiles([]);
    setSummary(null);
    setPointsPreview("");
    setImageMapPreview("");

    if (pointsInputRef.current) pointsInputRef.current.value = "";
    if (mapInputRef.current) mapInputRef.current.value = "";
    if (imagesInputRef.current) imagesInputRef.current.value = "";
  };

  const handlePointsFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    setPointsFile(file);
    setSummary(null);

    if (!file) {
      setPointsPreview("");
      return;
    }

    try {
      if (isExcelFile(file)) {
        const { headers, rows } = await readStructuredFile(file);
        const previewLines = [
          headers.join(", "),
          ...rows.slice(0, 5).map((r) => r.join(", ")),
        ];
        setPointsPreview(previewLines.join("\n"));
      } else {
        const text = await readTextFile(file);
        setPointsPreview(text.split(/\r?\n/).slice(0, 6).join("\n"));
      }
    } catch (err: any) {
      setPointsPreview("");
      alert(err?.message || String(err));
    }
  };

  const handleImageMapFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    setImageMapFile(file);
    setSummary(null);

    if (!file) {
      setImageMapPreview("");
      return;
    }

    try {
      if (isExcelFile(file)) {
        const { headers, rows } = await readStructuredFile(file);
        const previewLines = [
          headers.join(", "),
          ...rows.slice(0, 5).map((r) => r.join(", ")),
        ];
        setImageMapPreview(previewLines.join("\n"));
      } else {
        const text = await readTextFile(file);
        setImageMapPreview(text.split(/\r?\n/).slice(0, 6).join("\n"));
      }
    } catch (err: any) {
      setImageMapPreview("");
      alert(err?.message || String(err));
    }
  };

  const handleImagesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    setImageFiles(files);
    setSummary(null);
  };

  const runImportOptionA = async () => {
    if (!bulkProjectId) return alert("Select a project.");
    if (!bucketName.trim()) return alert("Bucket name is required.");
    if (!pointsFile) return alert("Select Points file.");
    if (!imageMapFile) return alert("Select Images Mapping file.");
    if (!imageFiles.length) return alert("Select image files (bulk).");

    setImporting(true);
    setSummary(null);

    const errors: string[] = [];

    try {
      const {
        data: { session },
        error: sessionErr,
      } = await supabase.auth.getSession();

      if (sessionErr) throw sessionErr;
      if (!session?.user) {
        redirectToLogin();
        return;
      }

      const [points, imageMap] = await Promise.all([
        parsePoints(pointsFile),
        parseImageMap(imageMapFile),
      ]);

      if (!points.length) throw new Error("No valid points found in Points file.");
      if (!imageMap.length) throw new Error("No valid rows found in Images Mapping file.");

      const duplicateSelectedImages = getDuplicates(imageFiles.map((f) => f.name));
      const duplicateMappingFiles = getDuplicates(imageMap.map((r) => r.file_name));
      const invalidCategories: string[] = [];

      if (duplicateSelectedImages.length) {
        errors.push(`Duplicate selected image names found: ${duplicateSelectedImages.join(", ")}`);
      }

      if (duplicateMappingFiles.length) {
        errors.push(`Duplicate file_name values found in mapping file: ${duplicateMappingFiles.join(", ")}`);
      }

      const pointByKey = new Map<string, ParsedPointRow>();
      points.forEach((p) => {
        pointByKey.set(p.point_key, {
          ...p,
          category: normalizeCategoryName(p.category),
        });
      });

      const missingPointKeysInPointsCsv = Array.from(
        new Set(
          imageMap
            .map((m) => m.point_key)
            .filter((pk): pk is string => !!pk && !pointByKey.has(pk))
        )
      ).sort((a, b) => {
        const na = Number(a);
        const nb = Number(b);
        if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
        return a.localeCompare(b);
      });

      if (missingPointKeysInPointsCsv.length) {
        throw new Error(
          `These point_key values are used in mapping file but missing in points file: ${missingPointKeysInPointsCsv.join(
            ", "
          )}`
        );
      }

      let noGpsReportId: string | null = null;

      async function getOrCreateNoGpsReport() {
        if (noGpsReportId) return noGpsReportId;

        const noGpsCategory = "NO_GPS Images";

        const { data: existing, error: exErr } = await supabase
          .from("reports")
          .select("id")
          .eq("project_id", bulkProjectId)
          .eq("category", noGpsCategory)
          .order("created_at", { ascending: false })
          .limit(1);

        if (exErr) throw exErr;

        if (existing && existing.length) {
          noGpsReportId = existing[0].id;
          return noGpsReportId;
        }

        const { data: created, error: crErr } = await supabase
          .from("reports")
          .insert([
            {
              project_id: bulkProjectId,
              category: noGpsCategory,
              description: "Images that do not have GPS point mapping (bulk import).",
              route_id: null,
              difficulty: "green",
            },
          ])
          .select("id")
          .single();

        if (crErr) throw crErr;

        noGpsReportId = created?.id ?? null;
        if (!noGpsReportId) throw new Error("Failed to create NO_GPS report.");

        return noGpsReportId;
      }

      const reportIdByPointKey = new Map<string, string>();
      const sortedKeys = Array.from(pointByKey.keys()).sort((a, b) => {
        const na = Number(a);
        const nb = Number(b);
        if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
        return a.localeCompare(b);
      });

      for (const key of sortedKeys) {
        const p = pointByKey.get(key)!;
        const category = p.category.trim() || "Unknown";
        const description = p.description && p.description.trim() ? p.description.trim() : null;

        const { data: found, error: fErr } = await supabase
          .from("reports")
          .select("id")
          .eq("project_id", bulkProjectId)
          .eq("category", category)
          .eq("description", description)
          .order("created_at", { ascending: false })
          .limit(1);

        if (fErr) throw fErr;

        let reportId: string | null = found && found.length ? found[0].id : null;

        if (!reportId) {
          const { data: created, error: cErr } = await supabase
            .from("reports")
            .insert([
              {
                project_id: bulkProjectId,
                category,
                description,
                route_id: null,
                difficulty: "green",
              },
            ])
            .select("id")
            .single();

          if (cErr) throw cErr;
          reportId = created?.id ?? null;
        }

        if (!reportId) throw new Error(`Failed to create/find report for point_key=${key}`);

        reportIdByPointKey.set(key, reportId);

        const { error: delErr } = await supabase
          .from("report_path_points")
          .delete()
          .eq("report_id", reportId);

        if (delErr) throw delErr;

        const { error: insErr } = await supabase
          .from("report_path_points")
          .insert([
            {
              report_id: reportId,
              seq: 1,
              latitude: p.latitude,
              longitude: p.longitude,
              elevation: null,
              accuracy: null,
              timestamp: null,
            },
          ]);

        if (insErr) throw insErr;
      }

      const mapByFileName = new Map<string, ParsedImageMapRow>();
      imageMap.forEach((r) => {
        if (!mapByFileName.has(r.file_name)) mapByFileName.set(r.file_name, r);
      });

      const selectedByName = new Map<string, File>();
      imageFiles.forEach((f) => {
        if (!selectedByName.has(f.name)) selectedByName.set(f.name, f);
      });

      const missingFilesInUpload: string[] = [];
      for (const [name] of mapByFileName) {
        if (!selectedByName.has(name)) missingFilesInUpload.push(name);
      }

      const extraFilesNotInMap: string[] = [];
      imageFiles.forEach((f) => {
        if (!mapByFileName.has(f.name)) extraFilesNotInMap.push(f.name);
      });

      const noGpsImages: string[] = [];
      let imagesUploaded = 0;
      let photosInserted = 0;

      await mapLimit(imageFiles, 3, async (file) => {
        const mapping = mapByFileName.get(file.name);
        const point_key = mapping?.point_key ?? null;

        let reportId = point_key ? reportIdByPointKey.get(point_key) : null;
        if (!reportId) {
          reportId = await getOrCreateNoGpsReport();
          noGpsImages.push(file.name);
        }

        const safeFileName = file.name.replace(/[^\w.\-]+/g, "_");
        const storagePath = `reports/${bulkProjectId}/${reportId}/${Date.now()}_${safeFileName}`;

        let uploaded: { path: string; publicUrl: string };
        try {
          uploaded = await uploadToBucket(bucketName.trim(), storagePath, file);
        } catch (upErr: any) {
          errors.push(`${file.name}: upload failed - ${upErr?.message || String(upErr)}`);
          return;
        }

        imagesUploaded++;

        const url = uploaded.publicUrl;

        if (!url) {
          errors.push(`${file.name}: could not generate public URL`);
          return;
        }

        const { width, height } = await getImageSize(file);

        const { error: insErr } = await supabase.from("report_photos").insert([
          {
            report_id: reportId,
            url,
            width,
            height,
          },
        ]);

        if (insErr) {
          errors.push(`${file.name}: report_photos insert failed - ${insErr.message}`);
          return;
        }

        photosInserted++;
      });

      setSummary({
        pointsRead: points.length,
        reportsCreatedOrUsed: sortedKeys.length + (noGpsReportId ? 1 : 0),
        imagesSelected: imageFiles.length,
        imagesUploaded,
        photosInserted,
        noGpsImages,
        missingFilesInUpload,
        extraFilesNotInMap,
        errors,
        duplicateSelectedImages,
        duplicateMappingFiles,
        missingPointKeysInPointsCsv,
        invalidCategories,
      });

      alert(
        errors.length
          ? "Bulk import completed with some warnings/errors. Check summary."
          : "Bulk import completed."
      );
    } catch (e: any) {
      const msg = String(e?.message || e || "").toLowerCase();
      if (msg.includes("auth") || msg.includes("session") || msg.includes("jwt")) {
        redirectToLogin();
        return;
      }
      alert(e?.message || String(e));
    } finally {
      setImporting(false);
    }
  };

  return (
    <div style={styles.page}>
      <div style={styles.header}>
        <div>
          <div style={styles.title}>Projects</div>
          <div style={styles.subtitle}>
            Total: <b>{projects.length}</b> • Showing: <b>{filtered.length}</b>
          </div>
        </div>

        <div style={styles.headerRight}>
          <button
            style={{ ...styles.btnGhost, opacity: loading || exporting ? 0.7 : 1 }}
            onClick={load}
            disabled={loading || exporting}
          >
            {loading ? "Refreshing..." : "Refresh"}
          </button>

          <button style={styles.btnPrimary} onClick={() => setNewOpen(true)}>
            + New Project
          </button>

          <button style={styles.btnGhost} onClick={() => setBulkOpen(true)}>
            Bulk Import (GPS + Images)
          </button>

          <div style={styles.exportGroup}>
            <button
              style={{ ...styles.btnPrimary, opacity: exporting ? 0.7 : 1 }}
              onClick={exportCSV}
              disabled={!filtered.length || exporting}
            >
              {exporting ? "Exporting..." : "Export CSV"}
            </button>
          </div>

          <button style={styles.btnDanger} onClick={logout}>
            Logout
          </button>
        </div>
      </div>

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

      {loading ? (
        <div style={styles.stateBox}>Loading projects...</div>
      ) : filtered.length === 0 ? (
        <div style={styles.stateBox}>No projects found</div>
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

      {newOpen && (
        <div style={styles.modalOverlay} onClick={() => setNewOpen(false)}>
          <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
            <div style={{ fontSize: 18, fontWeight: 800, marginBottom: 10 }}>
              Create Project
            </div>

            <div style={styles.formRow}>
              <div style={styles.formLabel}>Name *</div>
              <input
                style={styles.input}
                value={newName}
                onChange={(e) => setNewName(e.target.value)}
                placeholder="Project name"
              />
            </div>

            <div style={styles.formRow}>
              <div style={styles.formLabel}>Description</div>
              <textarea
                style={styles.textarea}
                value={newDesc}
                onChange={(e) => setNewDesc(e.target.value)}
                placeholder="Optional description"
              />
            </div>

            <div style={styles.modalActions}>
              <button style={styles.btnGhost} onClick={() => setNewOpen(false)}>
                Cancel
              </button>
              <button
                style={{ ...styles.btnPrimary, opacity: creating ? 0.7 : 1 }}
                onClick={createProject}
                disabled={creating}
              >
                {creating ? "Creating..." : "Create"}
              </button>
            </div>
          </div>
        </div>
      )}

      {bulkOpen && (
        <div style={styles.modalOverlay} onClick={() => setBulkOpen(false)}>
          <div style={styles.modalWide} onClick={(e) => e.stopPropagation()}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 10,
              }}
            >
              <div style={{ fontSize: 18, fontWeight: 800 }}>
                Bulk Import (GPS + Images)
              </div>
              <button style={styles.btnGhost} onClick={resetBulk} disabled={importing}>
                Clear
              </button>
            </div>

            <div style={{ marginTop: 10, fontSize: 13, color: "#475467", lineHeight: 1.55 }}>
              <b>Points File:</b> point_key, latitude, longitude, category (optional:
              description)
              <br />
              <b>Images File:</b> file_name, point_key (optional: image_key 6.1, 6.2...)
              <br />
              Supports <b>CSV, TXT, XLSX, XLS</b>.
              <br />
              Any category is allowed. Known categories are normalized automatically.
              <br />
              Files not mapped / NO_GPS → stored under <b>NO_GPS Images</b> report automatically.
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: 10,
                marginTop: 12,
              }}
            >
              <div>
                <div style={styles.formLabel}>Select Project</div>
                <select
                  style={styles.input as any}
                  value={bulkProjectId}
                  onChange={(e) => setBulkProjectId(e.target.value)}
                >
                  {projects.map((p) => (
                    <option key={p.id} value={p.id}>
                      {safeName(p)} ({p.id.slice(0, 8)}…)
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <div style={styles.formLabel}>Storage Bucket</div>
                <input
                  style={styles.input}
                  value={bucketName}
                  onChange={(e) => setBucketName(e.target.value)}
                />
              </div>
            </div>

            <div style={{ marginTop: 10 }}>
              <div style={styles.formLabel}>Category Handling</div>
              <div style={styles.categoryHelpBox}>
                Known categories are normalized automatically. New or custom categories are
                also allowed and will be imported as entered.
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: 10,
                marginTop: 12,
              }}
            >
              <div>
                <div style={styles.formLabel}>Points file</div>
                <input
                  ref={pointsInputRef}
                  type="file"
                  accept=".csv,.txt,.xlsx,.xls"
                  onChange={handlePointsFileChange}
                />
                {pointsFile && (
                  <div style={styles.fileMeta}>
                    {pointsFile.name} • {pointsFile.size} bytes
                  </div>
                )}
                {pointsPreview && <pre style={styles.previewBox}>{pointsPreview}</pre>}
              </div>

              <div>
                <div style={styles.formLabel}>Images Mapping file</div>
                <input
                  ref={mapInputRef}
                  type="file"
                  accept=".csv,.txt,.xlsx,.xls"
                  onChange={handleImageMapFileChange}
                />
                {imageMapFile && (
                  <div style={styles.fileMeta}>
                    {imageMapFile.name} • {imageMapFile.size} bytes
                  </div>
                )}
                {imageMapPreview && <pre style={styles.previewBox}>{imageMapPreview}</pre>}
              </div>
            </div>

            <div style={{ marginTop: 12 }}>
              <div style={styles.formLabel}>Select Images (bulk)</div>
              <input
                ref={imagesInputRef}
                type="file"
                multiple
                accept="image/*"
                onChange={handleImagesChange}
              />
              <div style={{ fontSize: 12, color: "#667085", marginTop: 6 }}>
                Selected: <b>{imageFiles.length}</b> images
              </div>
            </div>

            {summary && (
              <div style={{ ...styles.stateBox, marginTop: 12 }}>
                <div style={{ fontWeight: 800, marginBottom: 8 }}>Import Summary</div>

                <div style={{ fontSize: 13, lineHeight: 1.7 }}>
                  Points read: <b>{summary.pointsRead}</b>
                  <br />
                  Reports used: <b>{summary.reportsCreatedOrUsed}</b>
                  <br />
                  Images selected: <b>{summary.imagesSelected}</b>
                  <br />
                  Images uploaded: <b>{summary.imagesUploaded}</b>
                  <br />
                  report_photos inserted: <b>{summary.photosInserted}</b>
                </div>

                {summary.duplicateSelectedImages.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>Duplicate selected image names:</b>
                    <div style={styles.scrollBox}>
                      {summary.duplicateSelectedImages.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.duplicateMappingFiles.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>Duplicate mapping file_name values:</b>
                    <div style={styles.scrollBox}>
                      {summary.duplicateMappingFiles.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.missingPointKeysInPointsCsv.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>point_key in mapping file but missing in points file:</b>
                    <div style={styles.scrollBox}>
                      {summary.missingPointKeysInPointsCsv.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.noGpsImages.length > 0 && (
                  <div style={{ marginTop: 10, color: "#7A5AF8", fontSize: 13 }}>
                    <b>NO_GPS images:</b>
                    <div style={styles.scrollBox}>
                      {summary.noGpsImages.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.missingFilesInUpload.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>In mapping file but not selected:</b>
                    <div style={styles.scrollBox}>
                      {summary.missingFilesInUpload.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.extraFilesNotInMap.length > 0 && (
                  <div style={{ marginTop: 10, color: "#475467", fontSize: 13 }}>
                    <b>Selected images not in mapping (treated as NO_GPS):</b>
                    <div style={styles.scrollBox}>
                      {summary.extraFilesNotInMap.map((f) => (
                        <div key={f}>{f}</div>
                      ))}
                    </div>
                  </div>
                )}

                {summary.errors.length > 0 && (
                  <div style={{ marginTop: 10, color: "#B42318", fontSize: 13 }}>
                    <b>Errors / Warnings:</b>
                    <div style={styles.scrollBox}>
                      {summary.errors.map((e, idx) => (
                        <div key={idx}>{e}</div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            <div style={styles.modalActions}>
              <button
                style={styles.btnGhost}
                onClick={() => setBulkOpen(false)}
                disabled={importing}
              >
                Close
              </button>
              <button
                style={{ ...styles.btnPrimary, opacity: importing ? 0.7 : 1 }}
                onClick={runImportOptionA}
                disabled={importing}
              >
                {importing ? "Importing..." : "Run Import"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

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
  exportGroup: { display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" },
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
  searchInput: { flex: 1, border: "none", outline: "none", fontSize: 14, color: "#101828" },
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
  grid: { display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: 12 },
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
  cardTitle: { fontSize: 16, fontWeight: 800, lineHeight: 1.2 },
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
  cardHint: { marginTop: 10, fontSize: 12, color: "#475467", fontWeight: 700 },
  modalOverlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,0.35)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 16,
    zIndex: 9999,
  },
  modal: {
    width: "min(640px, 100%)",
    background: "#fff",
    borderRadius: 16,
    border: "1px solid #EAECF0",
    padding: 16,
    boxShadow: "0 10px 30px rgba(0,0,0,0.15)",
  },
  modalWide: {
    width: "min(980px, 100%)",
    background: "#fff",
    borderRadius: 16,
    border: "1px solid #EAECF0",
    padding: 16,
    boxShadow: "0 10px 30px rgba(0,0,0,0.15)",
    maxHeight: "90vh",
    overflow: "auto",
  },
  formRow: { marginTop: 10 },
  formLabel: { fontSize: 12, fontWeight: 800, color: "#344054", marginBottom: 6 },
  input: {
    width: "100%",
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #D0D5DD",
    outline: "none",
    fontSize: 14,
  },
  textarea: {
    width: "100%",
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid #D0D5DD",
    outline: "none",
    fontSize: 14,
    minHeight: 90,
    resize: "vertical",
  },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 14 },
  scrollBox: {
    maxHeight: 160,
    overflow: "auto",
    marginTop: 6,
    border: "1px solid #EAECF0",
    borderRadius: 10,
    padding: 8,
    background: "#fff",
  },
  fileMeta: { marginTop: 6, fontSize: 12, color: "#667085" },
  previewBox: {
    marginTop: 8,
    background: "#F8FAFC",
    border: "1px solid #EAECF0",
    borderRadius: 10,
    padding: 10,
    fontSize: 12,
    lineHeight: 1.5,
    color: "#101828",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    maxHeight: 140,
    overflow: "auto",
  },
  categoryHelpBox: {
    marginTop: 4,
    background: "#F8FAFC",
    border: "1px solid #EAECF0",
    borderRadius: 10,
    padding: 10,
    fontSize: 12,
    lineHeight: 1.6,
    color: "#101828",
  },
};