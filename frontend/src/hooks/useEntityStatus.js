import { useState, useEffect, useCallback } from "react";
import { API_ENDPOINTS } from "../config/api";

export function useEntityStatus() {
  const [status, setStatus]   = useState({});   // { Supplier: { files: [], count: 0 }, … }
  const [loading, setLoading] = useState(true);

  const refresh = useCallback(async () => {
    try {
      const res  = await fetch(API_ENDPOINTS.STATUS);
      const data = await res.json();
      setStatus(data);
    } catch (err) {
      console.error("Failed to fetch status:", err);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    refresh();
    const id = setInterval(refresh, 5000); // poll every 5s
    return () => clearInterval(id);
  }, [refresh]);

  return { status, loading, refresh };
}
