#! python3, this code written by kubikanber, 09.11.2018, author: kubikanber
# AcroExch.Point
# Defines the location of an AcroPoint.
#
# Properties
# The Point object has the following properties.
#
# Property
#
# Description
#
# X
#
# Gets or sets the x-coordinate of an AcroPoint.
#
# Y
#
# Gets or sets the y-coordinate of an AcroPoint.
from win32com.client.dynamic import Dispatch


class Point:

    def __init__(self, point=None):
        if point is None:
            self.point = Dispatch("AcroExch.Point")
        else:
            self.point = point

    def x(self, value=None):
        if value is not None:
            self.point.X = value
        return self.point.X

    def y(self, value=None):
        if value is not None:
            self.point.Y = value
        return self.point.Y
