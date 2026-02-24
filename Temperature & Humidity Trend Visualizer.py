import tkinter as tk
from tkinter import messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import random

class TempHumidityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Temperature & Humidity Trend Visualizer")
        self.root.geometry("900x600")

        self.temperatures = []
        self.humidities = []
        self.time_points = []

        self.create_widgets()
        self.create_plot()

    def create_widgets(self):
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10)

        tk.Label(input_frame, text="Temperature (°C):").grid(row=0, column=0, padx=5)
        self.temp_entry = tk.Entry(input_frame)
        self.temp_entry.grid(row=0, column=1, padx=5)

        tk.Label(input_frame, text="Humidity (%):").grid(row=0, column=2, padx=5)
        self.humidity_entry = tk.Entry(input_frame)
        self.humidity_entry.grid(row=0, column=3, padx=5)

        tk.Button(input_frame, text="Add Data", command=self.add_data).grid(row=0, column=4, padx=10)
        tk.Button(input_frame, text="Simulate Data", command=self.simulate_data).grid(row=0, column=5, padx=10)
        tk.Button(input_frame, text="Clear Data", command=self.clear_data).grid(row=0, column=6, padx=10)

    def create_plot(self):
        self.figure, self.ax = plt.subplots(figsize=(8,4))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.ax.set_title("Temperature & Humidity Trends")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Value")

    def add_data(self):
        try:
            temp = float(self.temp_entry.get())
            humidity = float(self.humidity_entry.get())

            self.temperatures.append(temp)
            self.humidities.append(humidity)
            self.time_points.append(len(self.time_points) + 1)

            self.update_plot()

            self.temp_entry.delete(0, tk.END)
            self.humidity_entry.delete(0, tk.END)

        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numeric values.")

    def simulate_data(self):
        temp = random.uniform(20, 35)
        humidity = random.uniform(40, 80)

        self.temperatures.append(round(temp, 2))
        self.humidities.append(round(humidity, 2))
        self.time_points.append(len(self.time_points) + 1)

        self.update_plot()

    def update_plot(self):
        self.ax.clear()
        self.ax.plot(self.time_points, self.temperatures, marker='o', label="Temperature (°C)")
        self.ax.plot(self.time_points, self.humidities, marker='s', label="Humidity (%)")

        self.ax.set_title("Temperature & Humidity Trends")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Value")
        self.ax.legend()
        self.ax.grid(True)

        self.canvas.draw()

    def clear_data(self):
        self.temperatures.clear()
        self.humidities.clear()
        self.time_points.clear()
        self.update_plot()


if __name__ == "__main__":
    root = tk.Tk()
    app = TempHumidityApp(root)
    root.mainloop()