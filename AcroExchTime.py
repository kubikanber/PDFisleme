#! python3, this code written by kubikanber, 10.11.2018, author: kubikanber
# AcroExch.Time
# Defines a specified time, accurate to the millisecond.
#
# Properties
# The Time object has the following properties.
#
# Property####### Description ##########################################################################################
#
# Date          # Gets or sets the date from an AcroTime.
#
# Hour          # Gets or sets the hour from an AcroTime.
#
# Millisecond   # Gets or sets the milliseconds from an AcroTime.
#
# Minute        # Gets or sets the minutes from an AcroTime.
#
# Month         # Gets or sets the month from an AcroTime.
#
# Second        # Gets or sets the seconds from an AcroTime.
#
# Year          # Gets or sets the year from an AcroTime.
########################################################################################################################
from win32com.client.dynamic import Dispatch


class AcroTime:

    def __init__(self, value=None):
        if value is not None:
            self.acrotime = value
        else:
            self.acrotime = Dispatch("AcroExch.Time")

    # Date
    # Gets or sets the date from an AcroTime.
    #
    # Syntax
    # [get/set] Short
    #
    # Returns
    # The date from the AcroTime. The date runs from 1 to 31.
    def date(self, value=None):
        if value is not None:
            self.acrotime.Date = value
        return self.acrotime.Date

    # Month
    # Gets or sets the month from an AcroTime.
    #
    # Syntax
    # [get/set] Short
    #
    # Returns
    # The month from the AcroTime. The month runs from 1 to 12, where 1 is January and 12 is December.
    def month(self, value=None):
        if value is not None:
            self.acrotime.Month = value
        return self.acrotime.Month

    # Year
    # Gets or sets the year from an AcroTime.
    #
    # Syntax
    # [get/set] Short
    #
    # Returns
    # The year from the AcroTime. The Year runs from 1 to 32767.
    def year(self, value=None):
        if value is not None:
            self.acrotime.Year = value
        return self.acrotime.Year

    # Hour
    # Gets or sets the hour from an AcroTime.
    #
    # Syntax
    # [get/set] Short
    #
    # Returns
    # The hour from the AcroTime. The hour runs from 0 to 23.
    def hour(self, value=None):
        if value is not None:
            self.acrotime.Hour = value
        return self.acrotime.Hour

    # Minute
    # Gets or sets the minutes from an AcroTime.
    #
    # Syntax
    # [get/set] Short
    #
    # Returns
    # The minutes from the AcroTime. Minutes run from 0 to 59.
    def minute(self, value=None):
        if value is not None:
            self.acrotime.Minute = value
        return self.acrotime.Minute

    # Second
    # Gets or sets the seconds from an AcroTime.
    #
    # Syntax
    # [get/set] Short
    #
    # Returns
    # The seconds from the AcroTime. Seconds run from 0 to 59.
    def second(self, value=None):
        if value is not None:
            self.acrotime.Second = value
        return self.acrotime.Second

    # Millisecond
    # Gets or sets the milliseconds from an AcroTime.
    #
    # Syntax
    # [get/set] Short
    #
    # Returns
    # The milliseconds from the AcroTime. Milliseconds run from 0 to 999.
    def millisecond(self, value=None):
        if value is not None:
            self.acrotime.Millisecond = value
        return self.acrotime.Millisecond
