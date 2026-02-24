import tkinter as tk
from tkinter import messagebox
import random
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class SensorAlertSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Real-Time Sensor Alert System")
        self.root.geometry("1000x650")

        self.sensor_values = []
        self.time_points = []
        self.running = False

        self.create_widgets()
        self.create_plot()

    def create_widgets(self):
        control_frame = tk.Frame(self.root)
        control_frame.pack(pady=10)

        tk.Label(control_frame, text="Upper Threshold:").grid(row=0, column=0, padx=5)
        self.upper_entry = tk.Entry(control_frame)
        self.upper_entry.grid(row=0, column=1, padx=5)
        self.upper_entry.insert(0, "75")

        tk.Label(control_frame, text="Lower Threshold:").grid(row=0, column=2, padx=5)
        self.lower_entry = tk.Entry(control_frame)
        self.lower_entry.grid(row=0, column=3, padx=5)
        self.lower_entry.insert(0, "25")

        tk.Button(control_frame, text="Start Monitoring", command=self.start_monitoring).grid(row=0, column=4, padx=10)
        tk.Button(control_frame, text="Stop", command=self.stop_monitoring).grid(row=0, column=5, padx=10)
        tk.Button(control_frame, text="Clear", command=self.clear_data).grid(row=0, column=6, padx=10)

        self.alert_label = tk.Label(self.root, text="", font=("Arial", 14))
        self.alert_label.pack(pady=10)

    def create_plot(self):
        self.figure, self.ax = plt.subplots(figsize=(9,4))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.ax.set_title("Real-Time Sensor Data")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Sensor Value")

    def start_monitoring(self):
        try:
            self.upper_threshold = float(self.upper_entry.get())
            self.lower_threshold = float(self.lower_entry.get())
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid threshold values.")
            return

        self.running = True
        self.update_sensor()

    def stop_monitoring(self):
        self.running = False

    def update_sensor(self):
        if not self.running:
            return

        value = random.uniform(10, 100)
        self.sensor_values.append(round(value, 2))
        self.time_points.append(len(self.time_points) + 1)

        self.check_alert(value)
        self.update_plot()

        self.root.after(1000, self.update_sensor)  # Update every second

    def check_alert(self, value):
        if value > self.upper_threshold:
            self.alert_label.config(text="⚠ ALERT: Value Above Upper Threshold!", fg="red")
        elif value < self.lower_threshold:
            self.alert_label.config(text="⚠ ALERT: Value Below Lower Threshold!", fg="orange")
        else:
            self.alert_label.config(text="Status: Normal", fg="green")

    def update_plot(self):
        self.ax.clear()
        self.ax.plot(self.time_points, self.sensor_values, marker='o')
        self.ax.axhline(self.upper_threshold, linestyle='--')
        self.ax.axhline(self.lower_threshold, linestyle='--')

        self.ax.set_title("Real-Time Sensor Data")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Sensor Value")
        self.ax.grid(True)

        self.canvas.draw()

    def clear_data(self):
        self.sensor_values.clear()
        self.time_points.clear()
        self.alert_label.config(text="")
        self.update_plot()

if __name__ == "__main__":
    root = tk.Tk()
    app = SensorAlertSystem(root)
    root.mainloop()