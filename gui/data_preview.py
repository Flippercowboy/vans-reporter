"""Data preview and editing window."""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Callable
import copy
from ..data.models import ProjectSummary, PersonHours


class DataPreviewWindow:
    """Window for previewing and editing calculated hours."""

    def __init__(self, parent, summary: ProjectSummary, callback: Callable):
        """
        Initialise the preview window.

        Args:
            parent: Parent window
            summary: ProjectSummary to display
            callback: Function to call with updated summary when confirmed
        """
        self.summary = copy.deepcopy(summary)  # Work with a copy
        self.callback = callback

        # Create window
        self.window = tk.Toplevel(parent)
        self.window.title("Preview & Edit Hours")
        self.window.geometry("800x600")

        self._create_widgets()
        self._populate_data()

    def _create_widgets(self):
        """Create all GUI widgets."""
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title_label = ttk.Label(
            main_frame,
            text="Preview & Edit Hours",
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Projects section
        projects_label = ttk.Label(
            main_frame,
            text="Projects:",
            font=("Arial", 12, "bold")
        )
        projects_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 5))

        # Projects treeview
        projects_frame = ttk.Frame(main_frame)
        projects_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))

        projects_scroll = ttk.Scrollbar(projects_frame, orient=tk.VERTICAL)
        self.projects_tree = ttk.Treeview(
            projects_frame,
            columns=('Hours',),
            height=5,
            yscrollcommand=projects_scroll.set
        )
        projects_scroll.config(command=self.projects_tree.yview)

        self.projects_tree.heading('#0', text='Project Name')
        self.projects_tree.heading('Hours', text='Total Hours')
        self.projects_tree.column('#0', width=500)
        self.projects_tree.column('Hours', width=150)

        self.projects_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        projects_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Team members section
        team_label = ttk.Label(
            main_frame,
            text="Team Members:",
            font=("Arial", 12, "bold")
        )
        team_label.grid(row=3, column=0, sticky=tk.W, pady=(10, 5))

        # Team treeview
        team_frame = ttk.Frame(main_frame)
        team_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))

        team_scroll = ttk.Scrollbar(team_frame, orient=tk.VERTICAL)
        self.team_tree = ttk.Treeview(
            team_frame,
            columns=('Total', 'Complete', 'Remaining'),
            height=8,
            yscrollcommand=team_scroll.set
        )
        team_scroll.config(command=self.team_tree.yview)

        self.team_tree.heading('#0', text='Name')
        self.team_tree.heading('Total', text='Total')
        self.team_tree.heading('Complete', text='Complete')
        self.team_tree.heading('Remaining', text='Remaining')
        self.team_tree.column('#0', width=350)
        self.team_tree.column('Total', width=100)
        self.team_tree.column('Complete', width=100)
        self.team_tree.column('Remaining', width=100)

        self.team_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        team_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Summary section
        summary_frame = ttk.LabelFrame(main_frame, text="Summary", padding="10")
        summary_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 20))

        self.summary_label = ttk.Label(summary_frame, text="", font=("Arial", 11))
        self.summary_label.grid(row=0, column=0, sticky=tk.W)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=(10, 0))

        edit_btn = ttk.Button(
            button_frame,
            text="Edit Hours",
            command=self._edit_hours
        )
        edit_btn.grid(row=0, column=0, padx=5)

        recalc_btn = ttk.Button(
            button_frame,
            text="Recalculate Totals",
            command=self._recalculate
        )
        recalc_btn.grid(row=0, column=1, padx=5)

        cancel_btn = ttk.Button(
            button_frame,
            text="Cancel",
            command=self.window.destroy
        )
        cancel_btn.grid(row=0, column=2, padx=5)

        confirm_btn = ttk.Button(
            button_frame,
            text="Confirm",
            command=self._confirm,
            style="Accent.TButton"
        )
        confirm_btn.grid(row=0, column=3, padx=5)

        # Configure grid weights
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(4, weight=2)
        projects_frame.columnconfigure(0, weight=1)
        projects_frame.rowconfigure(0, weight=1)
        team_frame.columnconfigure(0, weight=1)
        team_frame.rowconfigure(0, weight=1)

    def _populate_data(self):
        """Populate the treeviews with data."""
        # Clear existing items
        for item in self.projects_tree.get_children():
            self.projects_tree.delete(item)
        for item in self.team_tree.get_children():
            self.team_tree.delete(item)

        # Populate projects
        for project_name, hours in self.summary.projects.items():
            self.projects_tree.insert(
                '',
                'end',
                text=project_name,
                values=(f"{round(hours)}h",)
            )

        # Populate team members
        sorted_people = sorted(
            self.summary.people.items(),
            key=lambda x: x[1].total_hours,
            reverse=True
        )

        for person_name, person_hours in sorted_people:
            self.team_tree.insert(
                '',
                'end',
                text=person_name,
                values=(
                    f"{round(person_hours.total_hours)}h",
                    f"{round(person_hours.complete_hours)}h",
                    f"{round(person_hours.remaining_hours)}h"
                )
            )

        # Update summary
        self._update_summary()

    def _update_summary(self):
        """Update the summary label."""
        summary_text = (
            f"Total Department Hours: {round(self.summary.total_hours)}h  |  "
            f"Completed: {round(self.summary.complete_hours)}h  |  "
            f"Remaining: {round(self.summary.remaining_hours)}h"
        )
        self.summary_label.config(text=summary_text)

    def _edit_hours(self):
        """Open dialog to edit hours for selected item (project or team member)."""
        # Check if a project or team member is selected
        project_selected = self.projects_tree.selection()
        team_selected = self.team_tree.selection()

        if project_selected:
            self._edit_project_hours(project_selected[0])
        elif team_selected:
            self._edit_team_member_hours(team_selected[0])
        else:
            messagebox.showinfo("No Selection", "Please select a project or team member to edit.")

    def _edit_project_hours(self, item):
        """Edit hours for a selected project."""
        project_name = self.projects_tree.item(item, 'text')
        current_hours = self.summary.projects.get(project_name, 0)

        # Create edit dialog
        dialog = tk.Toplevel(self.window)
        dialog.title(f"Edit Hours: {project_name}")
        dialog.geometry("400x150")
        dialog.transient(self.window)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(frame, text=f"Project: {project_name}", font=("Arial", 11, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 15)
        )

        ttk.Label(frame, text="Total Hours:").grid(row=1, column=0, sticky=tk.W, pady=5)

        hours_var = tk.StringVar(value=str(round(current_hours)))
        hours_entry = ttk.Entry(frame, textvariable=hours_var, width=20)
        hours_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        hours_entry.focus()

        def save_changes():
            try:
                new_hours = float(hours_var.get())
                if new_hours < 0:
                    messagebox.showerror("Invalid Input", "Hours must be non-negative.")
                    return

                # Update the project hours proportionally
                old_hours = self.summary.projects[project_name]
                ratio = new_hours / old_hours if old_hours > 0 else 1

                # Update all people working on this project proportionally
                for person_name, person_hours in self.summary.people.items():
                    if project_name in person_hours.project_hours:
                        person_hours.project_hours[project_name] = round(person_hours.project_hours[project_name] * ratio, 1)
                        if project_name in person_hours.project_hours_complete:
                            person_hours.project_hours_complete[project_name] = round(person_hours.project_hours_complete[project_name] * ratio, 1)
                        if project_name in person_hours.project_hours_remaining:
                            person_hours.project_hours_remaining[project_name] = round(person_hours.project_hours_remaining[project_name] * ratio, 1)

                # Recalculate totals (this will rebuild project totals from people's data)
                self._recalculate()

                dialog.destroy()
                messagebox.showinfo("Success", "Project hours updated successfully!")

            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter a valid number.")

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(15, 0))

        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Save", command=save_changes).grid(row=0, column=1, padx=5)

    def _edit_team_member_hours(self, item):
        """Edit hours for a selected team member."""
        person_name = self.team_tree.item(item, 'text')
        person_hours = self.summary.people.get(person_name)

        if not person_hours:
            return

        # Create edit dialog
        dialog = tk.Toplevel(self.window)
        dialog.title(f"Edit Hours: {person_name}")
        dialog.geometry("500x400")
        dialog.transient(self.window)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(frame, text=f"Team Member: {person_name}", font=("Arial", 11, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 15)
        )

        ttk.Label(frame, text="Edit hours by project:", font=("Arial", 10)).grid(
            row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 10)
        )

        # Create entry fields for each project
        project_vars = {}
        row = 2
        for project_name, hours in person_hours.project_hours.items():
            ttk.Label(frame, text=f"{project_name}:").grid(row=row, column=0, sticky=tk.W, pady=5)
            var = tk.StringVar(value=str(round(hours)))
            entry = ttk.Entry(frame, textvariable=var, width=15)
            entry.grid(row=row, column=1, sticky=tk.W, pady=5)
            project_vars[project_name] = var
            row += 1

        def save_changes():
            try:
                # Update person's project hours
                for project_name, var in project_vars.items():
                    new_hours = float(var.get())
                    if new_hours < 0:
                        messagebox.showerror("Invalid Input", "Hours must be non-negative.")
                        return

                    old_hours = person_hours.project_hours[project_name]
                    ratio = new_hours / old_hours if old_hours > 0 else 1

                    person_hours.project_hours[project_name] = round(new_hours, 1)

                    # Update complete/remaining proportionally
                    if project_name in person_hours.project_hours_complete:
                        person_hours.project_hours_complete[project_name] = round(person_hours.project_hours_complete[project_name] * ratio, 1)
                    if project_name in person_hours.project_hours_remaining:
                        person_hours.project_hours_remaining[project_name] = round(person_hours.project_hours_remaining[project_name] * ratio, 1)

                # Recalculate totals
                self._recalculate()

                dialog.destroy()
                messagebox.showinfo("Success", "Team member hours updated successfully!")

            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter a valid number.")

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=(15, 0))

        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Save", command=save_changes).grid(row=0, column=1, padx=5)

    def _recalculate(self):
        """Recalculate all totals from the detailed data."""
        # Recalculate person totals from their project hours
        for person_name, person_hours in self.summary.people.items():
            person_hours.total_hours = round(sum(person_hours.project_hours.values()), 1)
            person_hours.complete_hours = round(sum(person_hours.project_hours_complete.values()), 1)
            person_hours.remaining_hours = round(sum(person_hours.project_hours_remaining.values()), 1)

        # Recalculate project totals from people's contributions
        for project_name in self.summary.projects.keys():
            project_total = 0
            project_complete = 0
            project_remaining = 0

            for person_hours in self.summary.people.values():
                project_total += person_hours.project_hours.get(project_name, 0)
                project_complete += person_hours.project_hours_complete.get(project_name, 0)
                project_remaining += person_hours.project_hours_remaining.get(project_name, 0)

            self.summary.projects[project_name] = round(project_total, 1)
            self.summary.projects_complete[project_name] = round(project_complete, 1)
            self.summary.projects_remaining[project_name] = round(project_remaining, 1)

        # Recalculate grand totals
        self.summary.total_hours = round(sum(self.summary.projects.values()), 1)
        self.summary.complete_hours = round(sum(self.summary.projects_complete.values()), 1)
        self.summary.remaining_hours = round(sum(self.summary.projects_remaining.values()), 1)

        # Refresh the display
        self._populate_data()

    def _confirm(self):
        """Confirm the data and close the window."""
        self.callback(self.summary)
        self.window.destroy()
