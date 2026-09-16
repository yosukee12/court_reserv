"""Place the initial window on the macOS display containing the pointer."""

import ctypes
import re
import sys


class Point(ctypes.Structure):
    _fields_ = [("x", ctypes.c_double), ("y", ctypes.c_double)]


class Size(ctypes.Structure):
    _fields_ = [("width", ctypes.c_double), ("height", ctypes.c_double)]


class Rect(ctypes.Structure):
    _fields_ = [("origin", Point), ("size", Size)]


def pointer_display_bounds():
    cg = ctypes.CDLL("/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics")
    cg.CGEventCreate.argtypes = [ctypes.c_void_p]
    cg.CGEventCreate.restype = ctypes.c_void_p
    cg.CGEventGetLocation.argtypes = [ctypes.c_void_p]
    cg.CGEventGetLocation.restype = Point
    cg.CGGetDisplaysWithPoint.argtypes = [
        Point, ctypes.c_uint32, ctypes.POINTER(ctypes.c_uint32),
        ctypes.POINTER(ctypes.c_uint32),
    ]
    cg.CGGetDisplaysWithPoint.restype = ctypes.c_int32
    cg.CGDisplayBounds.argtypes = [ctypes.c_uint32]
    cg.CGDisplayBounds.restype = Rect
    cf = ctypes.CDLL("/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation")
    cf.CFRelease.argtypes = [ctypes.c_void_p]
    cf.CFRelease.restype = None
    event = cg.CGEventCreate(None)
    if not event:
        raise RuntimeError("Pointer location unavailable")
    try:
        point = cg.CGEventGetLocation(event)
    finally:
        cf.CFRelease(event)
    display = ctypes.c_uint32()
    count = ctypes.c_uint32()
    status = cg.CGGetDisplaysWithPoint(point, 1, ctypes.byref(display), ctypes.byref(count))
    if status or not count.value:
        raise RuntimeError("Pointer display unavailable")
    bounds = cg.CGDisplayBounds(display)
    return (bounds.origin.x, bounds.origin.y, bounds.size.width, bounds.size.height)


def startup_geometry(saved, bounds):
    match = re.match(r"^(\d+)x(\d+)", saved)
    width, height = map(int, match.groups()) if match else (1180, 900)
    left, top, screen_width, screen_height = map(int, bounds)
    # Leave space for the menu bar, title bar and Dock.
    width = min(max(width, 1050), max(1, screen_width - 40))
    height = min(max(height, 820), max(1, screen_height - 100))
    x = left + (screen_width - width) // 2
    y = top + 35 + max(0, (screen_height - 100 - height) // 2)
    # A leading '+' keeps negative coordinates relative to the virtual origin.
    return f"{width}x{height}+{x}+{y}", (min(1050, width), min(820, height))


def position_startup_window(root, saved):
    bounds = (0, 0, root.winfo_screenwidth(), root.winfo_screenheight())
    if sys.platform == "darwin":
        try:
            bounds = pointer_display_bounds()
        except (OSError, RuntimeError, AttributeError):
            pass
    geometry, minimum = startup_geometry(saved, bounds)
    root.minsize(*minimum)
    root.geometry(geometry)
