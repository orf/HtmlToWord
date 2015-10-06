# -*- coding: utf-8 -*-
"""
helper module for styles
"""
import warnings

from win32com.client import constants


def getWdColorFromRGB(rgbAttr):
    """
    receive an rgb color attribute string like 'rgb(149, 55, 52)' and tranform it to a numeric constant
    in order to use it as a Selection.Font.Color attribute (as an item of WdColor enumeration)
    """
    try:
        values = rgbAttr[rgbAttr.find('(')+1:rgbAttr.find(')')].split(',')
    except:
        warnings.warn("getWdColorFromRGB: not possible to parse the RGB string '%s' " % rgbAttr)
        return None
    else:
        rgbstrlst = [v.strip() for v in values]
        return (int(rgbstrlst[0]) + 0x100 * int(rgbstrlst[1]) + 0x10000 * int(rgbstrlst[2]))

def getWdColorFromHex(hexAttr):
    """
    receive an hex color attribute string like '#9bbb59' (or '9bbb59') and tranform it to a numeric constant
    in order to use it as a Selection.Font.Color attribute (as an item of WdColor enumeration)
    """

    rgbstrlst = bytes.fromhex(hexAttr.strip('#'))
    return (int(rgbstrlst[0]) + 0x100 * int(rgbstrlst[1]) + 0x10000 * int(rgbstrlst[2]))


def getWdColorFromStyle(value):
    if value.split('(') and value.split('(')[0]=='rgb':
        return getWdColorFromRGB(value)
    else:
        return getWdColorFromHex(value)


def getWdColorIndexFromMapping(hex_value):
    """
    """
    hex_value = hex_value.strip('#')
    if hex_value in WORD_WDCOLORINDEX_MAPPING:
        return getattr(constants, WORD_WDCOLORINDEX_MAPPING.get(hex_value))
    return None



def getPointsFromPx(px_str):
    """
    receive an string representing the font-size attribute value in px (e.g. '16px') and tranform it
    to the equivalent value in points
    """
    if "pt" in px_str:
        return px_str.split("pt")[0]
    try:
        px = px_str.split('px')[0]
        return int(px)*0.75
    except (ValueError, IndexError):
        warnings.warn("Unable to tranform the value '%s' points" % px_str)
        return None

WORD_WDCOLORINDEX_MAPPING = {
    'ffff00': 'wdYellow',
    '00ff00': 'wdBrightGreen',
    '00ffff': 'wdTurquoise',
    'ff00ff': 'wdPink',
    '0000ff': 'wdBlue',
    'ff0000': 'wdRed',
    '000080': 'wdDarkBlue',
    '008080': 'wdTeal',
    '008000': 'wdGreen',
    '800080': 'wdViolet',
    '800000': 'wdDarkRed',
    '808000': 'wdDarkYellow',
    '808080': 'wdGray50',
    'c0c0c0': 'wdGray25',
    '000000': 'wdBlack'
}
