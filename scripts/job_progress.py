#!/usr/bin/env python3
"""Shared progress and cooperative-cancellation window for batch helpers."""

import os
import threading
import time


STATUSES = ("OK", "NOT_FOUND", "AMBIGUOUS", "ERROR", "CANCELLED")


def format_duration(seconds):
    seconds = max(0, int(round(seconds)))
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return "{}:{:02d}:{:02d}".format(hours, minutes, seconds)
    return "{}:{:02d}".format(minutes, seconds)


class ProgressReporter:
    """Thread-safe job state shared by the worker and optional Tk window."""

    def __init__(self, total):
        self.total = max(0, int(total))
        self.started = time.monotonic()
        self.cancel_event = threading.Event()
        self.done_event = threading.Event()
        self.lock = threading.Lock()
        self.completed = 0
        self.current = ""
        self.phase = "Starting..."
        self.counts = {status: 0 for status in STATUSES}

    def start_item(self, current, phase="Looking up..."):
        with self.lock:
            self.current = str(current or "")
            self.phase = phase

    def set_phase(self, phase):
        with self.lock:
            self.phase = str(phase or "")

    def finish_item(self, status):
        normalized = str(status or "ERROR").upper()
        if normalized not in self.counts:
            normalized = "ERROR"
        with self.lock:
            self.completed += 1
            self.counts[normalized] += 1

    def request_cancel(self):
        self.cancel_event.set()
        self.set_phase("Cancelling after the current request...")

    def cancelled(self):
        return self.cancel_event.is_set()

    def wait(self, seconds, phase="Waiting for provider rate limit..."):
        if seconds <= 0:
            return not self.cancelled()
        self.set_phase(phase)
        return not self.cancel_event.wait(seconds)

    def snapshot(self):
        with self.lock:
            return {
                "total": self.total,
                "completed": self.completed,
                "current": self.current,
                "phase": self.phase,
                "counts": dict(self.counts),
                "cancelled": self.cancelled(),
                "done": self.done_event.is_set(),
                "elapsed": time.monotonic() - self.started,
            }


def _run_without_window(reporter, worker):
    try:
        return worker(reporter)
    finally:
        reporter.done_event.set()


def run_with_progress(
    title,
    total,
    worker,
    enabled=True,
    always_on_top=True,
):
    """Run worker(reporter), showing a responsive progress/cancel window."""
    reporter = ProgressReporter(total)
    if not enabled or os.environ.get("MEDIACATALOG_NO_PROGRESS") == "1":
        return _run_without_window(reporter, worker)

    try:
        import tkinter as tk
        from tkinter import ttk

        root = tk.Tk()
    except Exception:
        return _run_without_window(reporter, worker)

    result_holder = {}
    error_holder = {}

    root.title(title)
    root.resizable(False, False)
    try:
        if always_on_top:
            root.attributes("-topmost", True)
    except Exception:
        pass

    container = ttk.Frame(root, padding=14)
    container.grid(row=0, column=0, sticky="nsew")

    heading_var = tk.StringVar(value="Starting...")
    current_var = tk.StringVar(value="")
    counts_var = tk.StringVar(value="")
    timing_var = tk.StringVar(value="Elapsed 0:00")

    ttk.Label(container, textvariable=heading_var, width=56).grid(
        row=0, column=0, sticky="w"
    )
    progress_bar = ttk.Progressbar(
        container, orient="horizontal", mode="determinate", length=420
    )
    progress_bar.grid(row=1, column=0, pady=(8, 8), sticky="ew")
    ttk.Label(container, textvariable=current_var, width=56).grid(
        row=2, column=0, sticky="w"
    )
    ttk.Label(container, textvariable=counts_var, width=56).grid(
        row=3, column=0, pady=(5, 0), sticky="w"
    )
    ttk.Label(container, textvariable=timing_var, width=56).grid(
        row=4, column=0, pady=(5, 0), sticky="w"
    )
    ttk.Label(
        container,
        text="Please leave the spreadsheet untouched while this window is open.",
        width=56,
    ).grid(row=5, column=0, pady=(10, 8), sticky="w")

    cancel_button = ttk.Button(
        container, text="Cancel", command=reporter.request_cancel
    )
    cancel_button.grid(row=6, column=0, sticky="e")

    def run_worker():
        try:
            result_holder["value"] = worker(reporter)
        except BaseException as exc:
            error_holder["error"] = exc
        finally:
            reporter.done_event.set()

    worker_thread = threading.Thread(target=run_worker, name="MediaCatalogJob")
    worker_thread.daemon = True
    worker_thread.start()

    def refresh():
        state = reporter.snapshot()
        total_count = state["total"]
        completed = state["completed"]
        progress_bar["maximum"] = max(1, total_count)
        progress_bar["value"] = completed

        heading_var.set(
            "{}  ({} of {})".format(state["phase"], completed, total_count)
        )
        current_var.set(
            "Current item: {}".format(state["current"])
            if state["current"]
            else ""
        )
        counts = state["counts"]
        counts_var.set(
            "Matched {}   Not found {}   Ambiguous {}   Errors {}".format(
                counts["OK"],
                counts["NOT_FOUND"],
                counts["AMBIGUOUS"],
                counts["ERROR"],
            )
        )

        elapsed = state["elapsed"]
        timing = "Elapsed {}".format(format_duration(elapsed))
        if completed > 0 and completed < total_count and not state["cancelled"]:
            remaining = (elapsed / completed) * (total_count - completed)
            timing += "   Estimated remaining {}".format(
                format_duration(remaining)
            )
        timing_var.set(timing)

        if state["cancelled"]:
            cancel_button.configure(state="disabled")

        if state["done"]:
            if error_holder:
                heading_var.set("The job stopped with an error.")
            elif state["cancelled"]:
                heading_var.set(
                    "Cancelled after {} of {} rows.".format(completed, total_count)
                )
            else:
                heading_var.set("Completed {} of {} rows.".format(completed, total_count))
            root.after(600, root.destroy)
            return

        root.after(200, refresh)

    root.protocol("WM_DELETE_WINDOW", reporter.request_cancel)
    root.after(100, refresh)
    root.mainloop()
    worker_thread.join()

    if error_holder:
        raise error_holder["error"]
    return result_holder.get("value")
