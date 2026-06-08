import { useCallback, useRef } from "react";
import { API_ENDPOINTS } from "../config/api";

/**
 * useSSEAction — fires a POST that returns an SSE stream.
 * Calls onEvent(parsed) for each SSE data line, onDone() when type==="done".
 */
export function useSSEAction() {
  const abortRef = useRef(null);

  const run = useCallback(async (endpoint, body, { onEvent, onDone, onError }) => {
    // Cancel any previous in-flight request
    if (abortRef.current) abortRef.current.abort();
    const controller = new AbortController();
    abortRef.current = controller;

    try {
      const res = await fetch(endpoint, {
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify(body),
        signal:  controller.signal,
      });

      if (!res.ok) throw new Error(`HTTP ${res.status}`);

      const reader  = res.body.getReader();
      const decoder = new TextDecoder();
      let   buffer  = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop(); // keep partial line

        for (const line of lines) {
          if (!line.startsWith("data: ")) continue;
          try {
            const evt = JSON.parse(line.slice(6));
            if (evt.type === "done") {
              onDone?.();
            } else {
              onEvent?.(evt);
            }
          } catch {
            // malformed JSON — skip
          }
        }
      }
    } catch (err) {
      if (err.name !== "AbortError") {
        onError?.(err.message || "Network error");
      }
    }
  }, []);

  const cancel = useCallback(() => {
    abortRef.current?.abort();
  }, []);

  return { run, cancel };
}
