# import sys
# import math
# from PyQt5.QtWidgets import *
# from PyQt5.QtGui import QPainter, QColor, QPen
# from PyQt5.QtCore import Qt, QTimer
# from PyQt5.QtCore import QPoint
#
# class AnimationWindow(QWidget):
#     def __init__(self, radius):
#         super().__init__()
#         self.radius = radius
#         self.angle = 0
#         self.position = QPoint(0, 0)
#         self.timer = QTimer(self)
#         self.timer.timeout.connect(self.update_animation)
#         self.timer.start(10)  # Update every 16 milliseconds (60 FPS)
#         self.showMaximized()
#         self.setWindowTitle('Animation')
#
#     def paintEvent(self, event):
#         painter = QPainter(self)
#         painter.setRenderHint(QPainter.Antialiasing)
#         painter.setBrush(QColor(255, 0, 0))
#         painter.drawRect(self.position.x(), self.position.y(), 50, 50)
#
#     def update_animation(self):
#         # Update the position of the square based on the current angle
#         self.position.setX(int(self.width() / 2 + self.radius * math.cos(math.radians(self.angle))))
#         self.position.setY(int(self.height() / 2 + self.radius * math.sin(math.radians(self.angle))))
#         # Increment the angle to move the square in a circular path
#         self.angle += 1
#         self.update()  # Redraw the window
#
# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     radius = 250  # Set the radius of the circular path here
#     window = AnimationWindow(radius)
#     window.show()
#     sys.exit(app.exec_())
#


import sys
import math
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPainter, QPixmap
from PyQt5.QtCore import Qt, QTimer, QPoint

class AnimationWindow(QWidget):
    def __init__(self, radius):
        super().__init__()
        self.setGeometry(100,100, 200, 200)
        self.radius = radius
        self.angle = 0
        self.position = QPoint(0, 0)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_animation)
        self.timer.start(10)  # Update every 16 milliseconds (60 FPS)
        self.showMaximized()
        self.setWindowTitle('Animation')
        self.image = QPixmap("addition/21.png")  # Replace "path_to_your_image.jpg" with the actual path to your image

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.drawPixmap(self.position.x(), self.position.y(), self.image)

    def update_animation(self):
        # Update the position of the image based on the current angle
        self.position.setX(int(self.width() / 2 + self.radius * math.cos(math.radians(self.angle))))
        self.position.setY(int(self.height() / 2 + self.radius * math.sin(math.radians(self.angle))))
        # Increment the angle to move the image in a circular path
        self.angle += 1
        self.update()  # Redraw the window

if __name__ == '__main__':
    app = QApplication(sys.argv)
    radius = 100  # Set the radius of the circular path here
    window = AnimationWindow(radius)
    window.show()
    sys.exit(app.exec_())