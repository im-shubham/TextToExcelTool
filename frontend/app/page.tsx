"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  DateEntry,
  MSG_EMPTY,
  Shift,
  SubmitOk,
  fetchDateEntries,
  fetchStatus,
  newSubmissionId,
  parseActivities,
  submitActivities,
} from "@/lib/api";

const DRAFT_KEY = "dwa-draft-v1";

type Draft = {
  date: string;
  shift: Shift | null;
  text: string;
  // The id of a submission that was sent but not confirmed; reused on retry so it is never saved twice.
  pending: { id: string; key: string } | null;
};

function todayLocal(): string {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function formatDate(iso: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso;
  const [y, m, d] = iso.split("-").map(Number);
  return new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" }).format(
    new Date(Date.UTC(y, m - 1, d)),
  );
}

function loadDraft(): Partial<Draft> {
  try {
    const raw = window.localStorage.getItem(DRAFT_KEY);
    return raw ? (JSON.parse(raw) as Partial<Draft>) : {};
  } catch {
    return {};
  }
}

function saveDraft(d: Draft) {
  try {
    window.localStorage.setItem(DRAFT_KEY, JSON.stringify(d));
  } catch {
    /* storage unavailable: the form still works, it just isn't remembered */
  }
}

function describeEntries(entries: DateEntry[]): string {
  return entries
    .map((e) => `${e.shift} Shift: ${e.activity_count} ${e.activity_count === 1 ? "activity" : "activities"}`)
    .join(", ");
}

export default function Page() {
  const [date, setDate] = useState(todayLocal);
  const [shift, setShift] = useState<Shift | null>(null);
  const [text, setText] = useState("");
  const [pending, setPending] = useState<Draft["pending"]>(null);
  const [loaded, setLoaded] = useState(false);

  const [submitting, setSubmitting] = useState(false);
  const [success, setSuccess] = useState<SubmitOk | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [confirmEntries, setConfirmEntries] = useState<DateEntry[] | null>(null);
  const [dateEntries, setDateEntries] = useState<DateEntry[] | null>(null);
  const [statusLine, setStatusLine] = useState<{ ok: boolean; text: string } | null>(null);
  const inFlight = useRef(false);

  const activities = useMemo(() => parseActivities(text), [text]);
  const payloadKey = useMemo(() => JSON.stringify([date, shift, activities]), [date, shift, activities]);

  // Restore an unsent draft (e.g. after a refresh or an error).
  useEffect(() => {
    const d = loadDraft();
    if (d.text) setText(d.text);
    if (d.text && d.date) setDate(d.date);
    if (d.shift === "A" || d.shift === "B") setShift(d.shift);
    if (d.pending) setPending(d.pending);
    setLoaded(true);
  }, []);

  useEffect(() => {
    if (loaded) saveDraft({ date, shift, text, pending });
  }, [loaded, date, shift, text, pending]);

  const refreshStatus = useCallback(async () => {
    const s = await fetchStatus();
    if (s.ok && "file_name" in s) {
      const latest = s.latest_entry
        ? `Latest entry: ${formatDate(s.latest_entry.date)}${s.latest_entry.shift ? ` · ${s.latest_entry.shift} Shift` : ""}`
        : "No entries yet";
      setStatusLine({ ok: true, text: `Master workbook: ${s.file_name} — ${latest}` });
    } else {
      setStatusLine({ ok: false, text: "message" in s ? s.message : "" });
    }
  }, []);

  const refreshDateEntries = useCallback(async (d: string) => {
    setDateEntries(d ? await fetchDateEntries(d) : null);
  }, []);

  useEffect(() => {
    refreshStatus();
  }, [refreshStatus]);

  useEffect(() => {
    refreshDateEntries(date);
  }, [date, refreshDateEntries]);

  async function send(confirmExisting: boolean) {
    if (inFlight.current) return; // ignore double clicks
    setError(null);
    setSuccess(null);
    if (!date) {
      setError("Please select the work date.");
      return;
    }
    if (!shift) {
      setError("Please select A Shift or B Shift.");
      return;
    }
    if (activities.length === 0) {
      setError(MSG_EMPTY);
      return;
    }

    const id = pending && pending.key === payloadKey ? pending.id : newSubmissionId();
    const nextPending = { id, key: payloadKey };
    setPending(nextPending);
    saveDraft({ date, shift, text, pending: nextPending });

    inFlight.current = true;
    setSubmitting(true);
    const res = await submitActivities({
      submission_id: id,
      date,
      shift,
      activities,
      confirm_existing: confirmExisting,
    });
    inFlight.current = false;
    setSubmitting(false);

    if (res.ok) {
      setSuccess(res.data);
      setConfirmEntries(null);
      setText("");
      setPending(null);
      saveDraft({ date, shift, text: "", pending: null });
      refreshStatus();
      refreshDateEntries(date);
      return;
    }
    if (res.error.code === "date_exists") {
      setConfirmEntries(res.error.entries ?? []);
      return;
    }
    if (res.error.code === "submission_conflict") {
      setPending(null);
    }
    setError(res.error.message);
  }

  const canSubmit = !submitting;

  return (
    <main className="mx-auto max-w-2xl px-4 py-6 sm:py-10">
      <header className="mb-6">
        <h1 className="text-2xl font-bold tracking-tight sm:text-3xl">Daily Work Activity</h1>
        <p className="mt-1 text-slate-600">
          Paste your activities and update the master Excel workbook automatically.
        </p>
        {statusLine && (
          <p className={`mt-2 text-xs ${statusLine.ok ? "text-slate-500" : "text-red-700"}`}>{statusLine.text}</p>
        )}
      </header>

      <section className="space-y-5 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-slate-200 sm:p-6">
        <div className="grid gap-5 sm:grid-cols-2">
          <div>
            <label htmlFor="date" className="mb-1 block text-sm font-medium">
              Work date
            </label>
            <input
              id="date"
              type="date"
              required
              value={date}
              onChange={(e) => {
                setDate(e.target.value);
                setConfirmEntries(null);
                setSuccess(null);
              }}
              className="w-full rounded-lg border border-slate-300 px-3 py-2.5 text-base focus:border-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-200"
            />
          </div>
          <fieldset>
            <legend className="mb-1 block text-sm font-medium">Shift</legend>
            <div className="grid grid-cols-2 gap-2" role="radiogroup" aria-label="Shift">
              {(["A", "B"] as const).map((s) => (
                <button
                  key={s}
                  type="button"
                  role="radio"
                  aria-checked={shift === s}
                  onClick={() => {
                    setShift(s);
                    setConfirmEntries(null);
                    setSuccess(null);
                  }}
                  className={`rounded-lg border px-3 py-2.5 text-base font-medium transition focus:outline-none focus:ring-2 focus:ring-blue-200 ${
                    shift === s
                      ? "border-blue-600 bg-blue-600 text-white"
                      : "border-slate-300 bg-white text-slate-700 hover:border-blue-400"
                  }`}
                >
                  {s} Shift
                </button>
              ))}
            </div>
          </fieldset>
        </div>

        {dateEntries && dateEntries.length > 0 && !confirmEntries && (
          <p className="rounded-lg bg-amber-50 px-3 py-2 text-sm text-amber-900 ring-1 ring-amber-200">
            {formatDate(date)} already has activities in the master workbook ({describeEntries(dateEntries)}).
          </p>
        )}

        <div>
          <label htmlFor="activities" className="mb-1 block text-sm font-medium">
            Work activities
          </label>
          <textarea
            id="activities"
            rows={9}
            value={text}
            onChange={(e) => {
              setText(e.target.value);
              setSuccess(null);
              setConfirmEntries(null);
            }}
            placeholder={
              "Paste your daily work activities here. Enter one activity per line.\n\nExample:\nMain line 4x scanner holders fixed\nPacking line 12x ESD chairs shifted\nToolroom material segregation completed"
            }
            className="w-full resize-y rounded-lg border border-slate-300 px-3 py-2.5 text-base leading-relaxed focus:border-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-200"
          />
        </div>

        <div className="rounded-xl bg-slate-50 p-4 ring-1 ring-slate-200">
          <h2 className="text-sm font-semibold uppercase tracking-wide text-slate-500">Preview</h2>
          <dl className="mt-2 grid grid-cols-3 gap-2 text-sm">
            <div>
              <dt className="text-slate-500">Date</dt>
              <dd className="font-medium">{date ? formatDate(date) : "—"}</dd>
            </div>
            <div>
              <dt className="text-slate-500">Shift</dt>
              <dd className="font-medium">{shift ? `${shift} Shift` : "Not selected"}</dd>
            </div>
            <div>
              <dt className="text-slate-500">Activities</dt>
              <dd className="font-medium">{activities.length}</dd>
            </div>
          </dl>
          {activities.length > 0 ? (
            <ol className="mt-3 list-decimal space-y-1 pl-5 text-sm">
              {activities.map((a, i) => (
                <li key={i} className="break-words">
                  {a}
                </li>
              ))}
            </ol>
          ) : (
            <p className="mt-3 text-sm text-slate-500">Your activities will appear here.</p>
          )}
        </div>

        {confirmEntries && (
          <div className="rounded-xl bg-amber-50 p-4 ring-1 ring-amber-300" role="alertdialog" aria-live="assertive">
            <p className="font-medium text-amber-900">Activities already exist for {formatDate(date)}.</p>
            {confirmEntries.length > 0 && (
              <p className="mt-1 text-sm text-amber-900">{describeEntries(confirmEntries)}</p>
            )}
            <p className="mt-1 text-sm text-amber-900">
              Do you want to add {activities.length === 1 ? "this activity" : `these ${activities.length} activities`} to{" "}
              {shift} Shift on this date?
            </p>
            <div className="mt-3 flex flex-col gap-2 sm:flex-row">
              <button
                type="button"
                onClick={() => send(true)}
                disabled={submitting}
                className="rounded-lg bg-amber-600 px-4 py-2.5 font-medium text-white hover:bg-amber-700 disabled:opacity-60"
              >
                {submitting ? "Adding…" : "Add to this date"}
              </button>
              <button
                type="button"
                onClick={() => setConfirmEntries(null)}
                disabled={submitting}
                className="rounded-lg border border-amber-400 bg-white px-4 py-2.5 font-medium text-amber-900 hover:bg-amber-100"
              >
                Cancel
              </button>
            </div>
          </div>
        )}

        {error && (
          <p className="rounded-lg bg-red-50 px-3 py-2.5 text-sm text-red-800 ring-1 ring-red-200" role="alert">
            {error}
          </p>
        )}

        {success && (
          <div className="rounded-lg bg-green-50 px-3 py-2.5 text-green-900 ring-1 ring-green-200" role="status">
            <p className="font-medium">{success.message}</p>
            <p className="mt-0.5 text-sm">
              {success.added} {success.added === 1 ? "activity" : "activities"} added for {formatDate(success.date)} ·{" "}
              {success.shift} Shift.
            </p>
          </div>
        )}

        {!confirmEntries && (
          <button
            type="button"
            onClick={() => send(false)}
            disabled={!canSubmit}
            className="w-full rounded-lg bg-blue-600 px-4 py-3 text-base font-semibold text-white shadow-sm hover:bg-blue-700 disabled:cursor-not-allowed disabled:bg-slate-400"
          >
            {submitting ? "Submitting…" : "Submit Activities"}
          </button>
        )}
      </section>

      <div className="mt-6 flex justify-center">
        <a
          href="/api/workbook/download"
          className="rounded-lg border border-slate-300 bg-white px-4 py-2.5 text-sm font-medium text-slate-700 shadow-sm hover:bg-slate-50"
        >
          Download Latest Excel
        </a>
      </div>
    </main>
  );
}
