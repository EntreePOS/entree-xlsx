import { executePlayground } from "./runtime.js";

self.addEventListener("message", async ({ data }) => {
  try {
    const result = await executePlayground(data.source);
    self.postMessage({ ok: true, ...result, bytes: result.bytes.buffer }, [result.bytes.buffer]);
  } catch (error) {
    self.postMessage({ ok: false, error: error instanceof Error ? error.message : String(error) });
  }
});
