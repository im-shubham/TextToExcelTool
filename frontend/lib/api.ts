export type Shift = "A" | "B";

export type DateEntry = { shift: string; activity_count: number };

export type SubmitOk = {
  status: "saved" | "duplicate";
  message: string;
  added: number;
  date: string;
  shift: Shift;
};

export type ApiError = {
  code: string;
  message: string;
  entries?: DateEntry[];
};

export const MSG_UNAVAILABLE =
  "Unable to access the master Excel workbook. No activities have been submitted. Please try again later.";
export const MSG_SAVE_FAILED =
  "The activities could not be saved. Please retry. Your submission has not been confirmed.";
export const MSG_EMPTY = "Please enter at least one work activity.";

/** One activity per line; blank lines are ignored and the wording is otherwise kept as typed. */
export function parseActivities(text: string): string[] {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
}

export function newSubmissionId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  // Fallback for non-secure contexts (plain http on a LAN address).
  const b = crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const h = Array.from(b, (x) => x.toString(16).padStart(2, "0")).join("");
  return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
}

export async function submitActivities(payload: {
  submission_id: string;
  date: string;
  shift: Shift;
  activities: string[];
  confirm_existing: boolean;
}): Promise<{ ok: true; data: SubmitOk } | { ok: false; error: ApiError }> {
  let res: Response;
  try {
    res = await fetch("/api/submissions", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      cache: "no-store",
    });
  } catch {
    // The request may or may not have reached the server; retrying with the same id is safe.
    return { ok: false, error: { code: "network", message: MSG_SAVE_FAILED } };
  }
  let body: unknown = null;
  try {
    body = await res.json();
  } catch {
    /* non-JSON error page from a proxy */
  }
  if (res.ok && body && typeof body === "object" && "status" in body) {
    return { ok: true, data: body as SubmitOk };
  }
  if (body && typeof body === "object" && "code" in body) {
    return { ok: false, error: body as ApiError };
  }
  const message = res.status === 502 || res.status === 503 || res.status === 504 ? MSG_UNAVAILABLE : MSG_SAVE_FAILED;
  return { ok: false, error: { code: `http_${res.status}`, message } };
}

export async function fetchDateEntries(date: string): Promise<DateEntry[] | null> {
  try {
    const res = await fetch(`/api/dates/${encodeURIComponent(date)}`, { cache: "no-store" });
    if (!res.ok) return null;
    return ((await res.json()) as { entries: DateEntry[] }).entries;
  } catch {
    return null;
  }
}

export type WorkbookStatus = {
  ok: boolean;
  file_name: string;
  latest_entry: { date: string; shift: string | null } | null;
};

export async function fetchStatus(): Promise<WorkbookStatus | { ok: false; message: string }> {
  try {
    const res = await fetch("/api/workbook/status", { cache: "no-store" });
    const body = await res.json();
    return res.ok ? (body as WorkbookStatus) : { ok: false, message: (body as ApiError).message ?? MSG_UNAVAILABLE };
  } catch {
    return { ok: false, message: MSG_UNAVAILABLE };
  }
}
