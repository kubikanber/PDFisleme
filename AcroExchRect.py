#! python3
# AcroExch.Rect
# Defines the location of an AcroRect.
#
# The Rect object has the following properties.
#
# Properties
# Property  # Description
#
# Bottom    # Gets or sets the bottom y-coordinate of an AcroRect.
#
# Left      # Gets or sets the left x-coordinate of an AcroRect.
#
# Right     # Gets or sets the right x-coordinate of an AcroRect.
#
# Top       # Gets or sets the top y-coordinate of an AcroRect.
from win32com.client.dynamic import Dispatch


class Rect:

    def __init__(self):
        self.rect = Dispatch("AcroExch.Rect")

    def bottom(self, value=None):
        if value is not None:
            self.rect.Bottom = value
        return self.rect.Bottom

    def left(self, value=None):
        if value is not None:
            self.rect.Left = value
        return self.rect.Left

    def right(self, value=None):
        if value is not None:
            self.rect.Right = value
        return self.rect.Right

    def top(self, value=None):
        if value is not None:
            self.rect.Top = value
        return self.rect.Top

    # def create_rect(self, alt_nokta=0, sol_nokta=0, sağ_nokta=72, üst_nokta=72):
    #     kare = self.rect
    #     kare.Bottom = alt_nokta
    #     kare.Left = sol_nokta
    #     kare.Right = sağ_nokta
    #     kare.Top = üst_nokta
    #     return kare

    def create_rect(self, bottom_point=0, left_point=0, right_point=72, top_point=72):
        self.bottom(bottom_point)
        self.left(left_point)
        self.right(right_point)
        self.top(top_point)
