"use client";

import React, { useEffect, useState } from "react";
import { usePathname, useRouter } from "next/navigation";
import { supabase } from "@/lib/supabaseClient";

export default function AuthGate({ children }: { children: React.ReactNode }) {
  const router = useRouter();
  const pathname = usePathname();
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let sub: any = null;

    const isPublicPath = (p: string) =>
      p.startsWith("/login") || p.startsWith("/auth");

    (async () => {
      const { data, error } = await supabase.auth.getSession();
      if (error) {
        console.error("getSession error:", error);
        setReady(true);
        return;
      }

      const session = data?.session;

      if (!session && pathname && !isPublicPath(pathname)) {
        router.replace("/login");
        setReady(true);
        return;
      }

      setReady(true);

      const { data: listener } = supabase.auth.onAuthStateChange((_event, s) => {
        const p = pathname || "";
        if (!s && !isPublicPath(p)) router.replace("/login");
      });

      sub = listener?.subscription;
    })();

    return () => {
      try {
        sub?.unsubscribe?.();
      } catch {}
    };
  }, [router, pathname]);

  if (!ready) return <div style={{ padding: 24 }}>Loading...</div>;
  return <>{children}</>;
}
