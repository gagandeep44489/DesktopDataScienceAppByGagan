import sys
from pathlib import Path
import pandas as pd

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QMessageBox
)

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


class TimeSeriesDashboard(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Multi-Source Time Series Dashboard")
        self.setGeometry(200, 100, 900, 600)

        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.setCentralWidget(self.canvas)

        self.load_and_plot_data()

    def load_and_plot_data(self):
        try:
            base_dir = Path(__file__).parent

            source1_path = base_dir / "source1.csv"
            source2_path = base_dir / "source2.csv"

            df1 = pd.read_csv(source1_path, parse_dates=["date"])
            df2 = pd.read_csv(source2_path, parse_dates=["date"])

            df1.sort_values("date", inplace=True)
            df2.sort_values("date", inplace=True)

            ax = self.figure.add_subplot(111)
            ax.clear()

            ax.plot(df1["date"], df1["value"], marker="o", label="Source 1")
            ax.plot(df2["date"], df2["value"], marker="o", label="Source 2")

            ax.set_title("Multi-Source Time Series")
            ax.set_xlabel("Date")
            ax.set_ylabel("Value")
            ax.legend()
            ax.grid(True)

            self.canvas.draw()

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


def main():
    app = QApplication(sys.argv)
    window = TimeSeriesDashboard()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
