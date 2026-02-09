import sys
import random
from collections import deque

from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import QTimer

import pyqtgraph as pg


class RealTimeSensorMonitor(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Real-Time Sensor Data Monitor")
        self.setGeometry(200, 100, 900, 500)

        self.plot_widget = pg.PlotWidget(title="Live Sensor Data")
        self.setCentralWidget(self.plot_widget)

        self.data_buffer = deque(maxlen=100)
        self.x = list(range(100))

        self.curve = self.plot_widget.plot(self.x, [0] * 100, pen='g')

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_data)
        self.timer.start(200)  # update every 200 ms

    def update_data(self):
        new_value = random.uniform(20, 30)  # simulated sensor value
        self.data_buffer.append(new_value)

        y = list(self.data_buffer)
        if len(y) < 100:
            y = [0] * (100 - len(y)) + y

        self.curve.setData(self.x, y)


def main():
    app = QApplication(sys.argv)
    window = RealTimeSensorMonitor()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
