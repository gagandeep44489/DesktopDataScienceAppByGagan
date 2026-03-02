import random
import tkinter as tk
from tkinter import messagebox

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


class StreamingDataVisualizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Streaming Data Visualizer")
        self.root.geometry("900x620")

        self.max_points = 120
        self.time_data = []
        self.value_data = []
        self.tick = 0
        self.is_streaming = False
        self.after_id = None

        self._build_ui()
        self._set_status("Ready")

    def _build_ui(self):
        tk.Label(
            self.root,
            text="Streaming Data Visualizer",
            font=("Arial", 18, "bold")
        ).pack(pady=10)

        controls = tk.Frame(self.root)
        controls.pack(pady=5)

        tk.Label(controls, text="Interval (ms):").grid(row=0, column=0, padx=5)
        self.interval_var = tk.StringVar(value="300")
        tk.Entry(controls, textvariable=self.interval_var, width=8).grid(
            row=0, column=1, padx=5
        )

        tk.Label(controls, text="Min value:").grid(row=0, column=2, padx=5)
        self.min_var = tk.StringVar(value="20")
        tk.Entry(controls, textvariable=self.min_var, width=8).grid(
            row=0, column=3, padx=5
        )

        tk.Label(controls, text="Max value:").grid(row=0, column=4, padx=5)
        self.max_var = tk.StringVar(value="80")
        tk.Entry(controls, textvariable=self.max_var, width=8).grid(
            row=0, column=5, padx=5
        )

        buttons = tk.Frame(self.root)
        buttons.pack(pady=5)

        tk.Button(
            buttons,
            text="Start Stream",
            command=self.start_stream,
            bg="green",
            fg="white",
            width=14
        ).grid(row=0, column=0, padx=6)

        tk.Button(
            buttons,
            text="Pause",
            command=self.pause_stream,
            bg="orange",
            fg="white",
            width=10
        ).grid(row=0, column=1, padx=6)

        tk.Button(
            buttons,
            text="Reset",
            command=self.reset_stream,
            bg="red",
            fg="white",
            width=10
        ).grid(row=0, column=2, padx=6)

        figure = Figure(figsize=(9, 4.8), dpi=100)
        self.ax = figure.add_subplot(111)
        self.ax.set_title("Live Streaming Data")
        self.ax.set_xlabel("Time Step")
        self.ax.set_ylabel("Value")
        self.ax.grid(True, alpha=0.3)

        (self.line,) = self.ax.plot([], [], color="royalblue", linewidth=2)

        self.canvas = FigureCanvasTkAgg(figure, master=self.root)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.status_var = tk.StringVar()
        tk.Label(
            self.root,
            textvariable=self.status_var,
            anchor="w",
            relief=tk.SUNKEN,
            padx=8
        ).pack(fill=tk.X, side=tk.BOTTOM)

    def _set_status(self, text):
        self.status_var.set(f"Status: {text}")

    def _validate_inputs(self):
        try:
            interval = int(self.interval_var.get())
            min_value = float(self.min_var.get())
            max_value = float(self.max_var.get())
        except ValueError:
            raise ValueError("Interval must be an integer and min/max must be numbers.")

        if interval < 50:
            raise ValueError("Interval must be at least 50 ms.")
        if min_value >= max_value:
            raise ValueError("Min value must be smaller than max value.")

        return interval, min_value, max_value

    def start_stream(self):
        if self.is_streaming:
            return

        try:
            interval, min_value, max_value = self._validate_inputs()
        except ValueError as exc:
            messagebox.showerror("Invalid Input", str(exc))
            return

        self.interval = interval
        self.min_value = min_value
        self.max_value = max_value
        self.is_streaming = True
        self._set_status("Streaming...")
        self._stream_step()

    def pause_stream(self):
        self.is_streaming = False
        if self.after_id is not None:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        self._set_status("Paused")

    def reset_stream(self):
        self.pause_stream()
        self.time_data.clear()
        self.value_data.clear()
        self.tick = 0
        self._update_plot()
        self._set_status("Reset complete")

    def _stream_step(self):
        if not self.is_streaming:
            return

        self.tick += 1
        new_value = random.uniform(self.min_value, self.max_value)

        self.time_data.append(self.tick)
        self.value_data.append(new_value)

        if len(self.time_data) > self.max_points:
            self.time_data = self.time_data[-self.max_points:]
            self.value_data = self.value_data[-self.max_points:]

        self._update_plot()
        self.after_id = self.root.after(self.interval, self._stream_step)

    def _update_plot(self):
        self.line.set_data(self.time_data, self.value_data)

        if self.time_data:
            self.ax.set_xlim(self.time_data[0], self.time_data[-1] + 1)

            min_y = min(self.value_data)
            max_y = max(self.value_data)
            padding = max((max_y - min_y) * 0.15, 1)
            self.ax.set_ylim(min_y - padding, max_y + padding)
        else:
            self.ax.set_xlim(0, 10)
            self.ax.set_ylim(0, 100)

        self.canvas.draw_idle()


def main():
    root = tk.Tk()
    app = StreamingDataVisualizer(root)
    def on_close():
        app.pause_stream()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()