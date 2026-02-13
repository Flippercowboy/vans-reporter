"""Main GUI window for Vans reporting application."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import date, timedelta
import calendar
import os
from ..data.monday_client import MondayClient
from ..data.calculator import TimeCalculator
from ..data.models import ProjectSummary
from ..powerpoint.generator import VansReportGenerator
from .data_preview import DataPreviewWindow


class MainWindow:
    """Main application window."""

    def __init__(self, root):
        """
        Initialise the main window.

        Args:
            root: Tk root window
        """
        self.root = root
        self.root.title("Vans Department Monthly Report")
        self.root.geometry("600x700")

        # Data storage
        self.projects = None
        self.summary = None
        self.as_of_date = date.today()

        # Initialise clients
        self.monday_client = MondayClient()
        self.calculator = TimeCalculator()

        self._create_widgets()
        self._update_date_display()

    def _create_widgets(self):
        """Create all GUI widgets."""
        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title_label = ttk.Label(
            main_frame,
            text="Vans Department Monthly Report",
            font=("Arial", 18, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Date selection
        date_frame = ttk.LabelFrame(main_frame, text="Report Settings", padding="10")
        date_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        # Month selector
        ttk.Label(date_frame, text="Report Month:", font=("Arial", 10)).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )

        month_year_frame = ttk.Frame(date_frame)
        month_year_frame.grid(row=0, column=1, sticky=tk.W, pady=5)

        # Month dropdown
        self.month_var = tk.StringVar(value=str(self.as_of_date.month))
        month_names = [
            "1 - January", "2 - February", "3 - March", "4 - April",
            "5 - May", "6 - June", "7 - July", "8 - August",
            "9 - September", "10 - October", "11 - November", "12 - December"
        ]
        month_combo = ttk.Combobox(
            month_year_frame,
            textvariable=self.month_var,
            values=month_names,
            width=15,
            state="readonly"
        )
        month_combo.set(f"{self.as_of_date.month} - {calendar.month_name[self.as_of_date.month]}")
        month_combo.grid(row=0, column=0, padx=(0, 5))
        month_combo.bind('<<ComboboxSelected>>', self._on_month_changed)

        # Year dropdown
        self.year_var = tk.StringVar(value=str(self.as_of_date.year))
        current_year = self.as_of_date.year
        years = [str(y) for y in range(current_year - 1, current_year + 2)]
        year_combo = ttk.Combobox(
            month_year_frame,
            textvariable=self.year_var,
            values=years,
            width=8,
            state="readonly"
        )
        year_combo.set(str(self.as_of_date.year))
        year_combo.grid(row=0, column=1)
        year_combo.bind('<<ComboboxSelected>>', self._on_month_changed)

        # As of date selector
        ttk.Label(date_frame, text="As of Date:", font=("Arial", 10)).grid(
            row=1, column=0, sticky=tk.W, pady=5
        )

        date_selector_frame = ttk.Frame(date_frame)
        date_selector_frame.grid(row=1, column=1, sticky=tk.W, pady=5)

        self.day_var = tk.StringVar(value=str(self.as_of_date.day))
        self.day_spinbox = ttk.Spinbox(
            date_selector_frame,
            from_=1,
            to=31,
            textvariable=self.day_var,
            width=5,
            command=self._on_day_changed
        )
        self.day_spinbox.grid(row=0, column=0, padx=(0, 10))

        # Today button
        today_btn = ttk.Button(date_selector_frame, text="Today", command=self._set_today)
        today_btn.grid(row=0, column=1, padx=(0, 5))

        # End of month button
        eom_btn = ttk.Button(date_selector_frame, text="End of Month", command=self._set_end_of_month)
        eom_btn.grid(row=0, column=2)

        # Date preview
        self.date_preview_label = ttk.Label(date_frame, text="", font=("Arial", 10, "italic"))
        self.date_preview_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))

        self._update_date_display()

        # Step 1: Fetch Data
        step1_frame = ttk.LabelFrame(main_frame, text="Step 1: Fetch Data from Monday.com", padding="15")
        step1_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        fetch_label = ttk.Label(
            step1_frame,
            text="Click below to fetch Vans department projects from Monday.com",
            wraplength=500
        )
        fetch_label.grid(row=0, column=0, pady=(0, 10))

        self.fetch_btn = ttk.Button(
            step1_frame,
            text="Fetch Vans Projects",
            command=self._fetch_projects,
            width=25
        )
        self.fetch_btn.grid(row=1, column=0)

        # Step 2: Preview Data
        step2_frame = ttk.LabelFrame(main_frame, text="Step 2: Review & Edit Hours", padding="15")
        step2_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        preview_label = ttk.Label(
            step2_frame,
            text="Review calculated hours and make adjustments if needed",
            wraplength=500
        )
        preview_label.grid(row=0, column=0, pady=(0, 10))

        self.preview_btn = ttk.Button(
            step2_frame,
            text="Preview Data",
            command=self._preview_data,
            width=25,
            state=tk.DISABLED
        )
        self.preview_btn.grid(row=1, column=0)

        # Step 3: Generate PowerPoint
        step3_frame = ttk.LabelFrame(main_frame, text="Step 3: Generate PowerPoint", padding="15")
        step3_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        save_frame = ttk.Frame(step3_frame)
        save_frame.grid(row=0, column=0, pady=(0, 10))

        ttk.Label(save_frame, text="Save to:").grid(row=0, column=0, sticky=tk.W)

        self.save_path_var = tk.StringVar()
        # Default save location
        default_path = os.path.expanduser("~/Desktop/Vans_Report.pptx")
        self.save_path_var.set(default_path)

        save_entry = ttk.Entry(save_frame, textvariable=self.save_path_var, width=40)
        save_entry.grid(row=0, column=1, padx=(10, 5))

        browse_btn = ttk.Button(save_frame, text="Browse", command=self._browse_save_location)
        browse_btn.grid(row=0, column=2)

        self.generate_btn = ttk.Button(
            step3_frame,
            text="Generate Presentation",
            command=self._generate_presentation,
            width=25,
            state=tk.DISABLED
        )
        self.generate_btn.grid(row=1, column=0, pady=(10, 0))

        # Status bar
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(20, 0))

        ttk.Label(status_frame, text="Status:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W)
        self.status_label = ttk.Label(status_frame, text="Ready", font=("Arial", 10))
        self.status_label.grid(row=0, column=1, sticky=tk.W, padx=(5, 0))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

    def _update_date_display(self):
        """Update the date display labels."""
        month_name = calendar.month_name[self.as_of_date.month]
        year = self.as_of_date.year
        day_name = calendar.day_name[self.as_of_date.weekday()]

        self.date_preview_label.config(
            text=f"Report will cover: {month_name} {year} (as of {day_name}, {self.as_of_date.day} {month_name} {year})"
        )

    def _on_month_changed(self, event=None):
        """Handle month/year selection change."""
        try:
            # Extract month number from "1 - January" format
            month_str = self.month_var.get().split(' - ')[0]
            month = int(month_str)
            year = int(self.year_var.get())

            # Keep the same day if valid, otherwise use last valid day of new month
            day = int(self.day_var.get())

            # Get the last day of the selected month
            if month == 12:
                next_month = date(year + 1, 1, 1)
            else:
                next_month = date(year, month + 1, 1)
            last_day_of_month = (next_month - timedelta(days=1)).day

            # Clamp day to valid range
            day = min(day, last_day_of_month)
            self.day_var.set(str(day))

            # Update spinbox range
            self.day_spinbox.config(to=last_day_of_month)

            self.as_of_date = date(year, month, day)
            self._update_date_display()
            self.status_label.config(text="Date updated")

            # Clear previous data when month changes
            self.projects = None
            self.summary = None
            self.preview_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.DISABLED)

        except ValueError:
            pass

    def _on_day_changed(self):
        """Handle day selection change."""
        try:
            month_str = self.month_var.get().split(' - ')[0]
            month = int(month_str)
            year = int(self.year_var.get())
            day = int(self.day_var.get())

            self.as_of_date = date(year, month, day)
            self._update_date_display()
            self.status_label.config(text="Date updated")

        except ValueError:
            pass

    def _set_today(self):
        """Set date to today."""
        today = date.today()
        self.as_of_date = today

        # Update all controls
        self.month_var.set(f"{today.month} - {calendar.month_name[today.month]}")
        self.year_var.set(str(today.year))
        self.day_var.set(str(today.day))

        # Update spinbox range for current month
        if today.month == 12:
            next_month = date(today.year + 1, 1, 1)
        else:
            next_month = date(today.year, today.month + 1, 1)
        last_day = (next_month - timedelta(days=1)).day
        self.day_spinbox.config(to=last_day)

        self._update_date_display()
        self.status_label.config(text="Set to today")

        # Clear previous data
        self.projects = None
        self.summary = None
        self.preview_btn.config(state=tk.DISABLED)
        self.generate_btn.config(state=tk.DISABLED)

    def _set_end_of_month(self):
        """Set date to end of selected month."""
        try:
            month_str = self.month_var.get().split(' - ')[0]
            month = int(month_str)
            year = int(self.year_var.get())

            # Get last day of month
            if month == 12:
                next_month = date(year + 1, 1, 1)
            else:
                next_month = date(year, month + 1, 1)
            last_day = (next_month - timedelta(days=1)).day

            self.day_var.set(str(last_day))
            self.as_of_date = date(year, month, last_day)
            self._update_date_display()
            self.status_label.config(text="Set to end of month")

            # Clear previous data
            self.projects = None
            self.summary = None
            self.preview_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.DISABLED)

        except ValueError:
            pass

    def _fetch_projects(self):
        """Fetch projects from Monday.com."""
        self.status_label.config(text="Fetching projects from Monday.com...")
        self.fetch_btn.config(state=tk.DISABLED)
        self.root.update()

        try:
            # Fetch projects
            self.projects = self.monday_client.get_vans_projects()

            if not self.projects:
                messagebox.showwarning(
                    "No Projects Found",
                    "No Vans department projects were found."
                )
                self.status_label.config(text="No projects found")
                self.fetch_btn.config(state=tk.NORMAL)
                return

            # Calculate hours
            self.summary = self.calculator.calculate_project_hours(self.projects, self.as_of_date)

            self.status_label.config(text=f"Fetched {len(self.projects)} projects successfully")
            self.preview_btn.config(state=tk.NORMAL)
            messagebox.showinfo(
                "Success",
                f"Fetched {len(self.projects)} Vans projects successfully!\n\n"
                f"Total hours: {int(self.summary.total_hours)}h\n"
                f"Completed: {int(self.summary.complete_hours)}h\n"
                f"Remaining: {int(self.summary.remaining_hours)}h"
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch projects:\n\n{str(e)}")
            self.status_label.config(text="Error fetching projects")
        finally:
            self.fetch_btn.config(state=tk.NORMAL)

    def _preview_data(self):
        """Open the data preview window."""
        if not self.summary:
            return

        preview_window = DataPreviewWindow(self.root, self.summary, self._on_data_confirmed)

    def _on_data_confirmed(self, updated_summary: ProjectSummary):
        """
        Callback when data is confirmed in preview window.

        Args:
            updated_summary: Updated ProjectSummary from preview window
        """
        self.summary = updated_summary
        self.status_label.config(text="Data confirmed, ready to generate")
        self.generate_btn.config(state=tk.NORMAL)

    def _browse_save_location(self):
        """Open file browser to select save location."""
        filename = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")],
            initialfile="Vans_Report.pptx"
        )
        if filename:
            self.save_path_var.set(filename)

    def _generate_presentation(self):
        """Generate the PowerPoint presentation."""
        if not self.summary:
            return

        self.status_label.config(text="Generating PowerPoint...")
        self.generate_btn.config(state=tk.DISABLED)
        self.root.update()

        try:
            output_path = self.save_path_var.get()

            # Generate presentation
            generator = VansReportGenerator(self.summary)
            generator.generate(output_path)

            self.status_label.config(text="Presentation generated successfully")
            messagebox.showinfo(
                "Success",
                f"Presentation generated successfully!\n\nSaved to:\n{output_path}"
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate presentation:\n\n{str(e)}")
            self.status_label.config(text="Error generating presentation")
        finally:
            self.generate_btn.config(state=tk.NORMAL)


def launch_gui():
    """Launch the GUI application."""
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
