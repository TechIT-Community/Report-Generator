"""
Utility functions for the Report Generator backend.
"""
from win32com.client import constants as c

def cm_to_pt(cm: float) -> float:
    """
    Converts centimeters to points.
    
    :param cm: The length in centimeters.
    :type cm: float
    :return: The length in points.
    :rtype: float
    """
    return cm * 28.35
