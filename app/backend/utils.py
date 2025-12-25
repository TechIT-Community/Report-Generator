"""
Utility functions for the Report Generator backend.
Contains helper conversions and shared constants.
"""

# =================================================================================================
#                                      UNIT CONVERSIONS
# =================================================================================================

def cm_to_pt(cm: float) -> float:
    """
    Converts centimeters to PostScript points.
    Used for specifying Word dimensions (margins, widths).
    
    Formula: 1 cm ≈ 28.35 points (72 points / 2.54 cm).
    
    :param cm: The length in centimeters.
    :return: The length in points.
    """
    return cm * 28.35
