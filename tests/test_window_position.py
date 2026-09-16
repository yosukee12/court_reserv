import unittest
from unittest.mock import Mock, patch

from court_reserv.ui.window_position import startup_geometry, position_startup_window


class WindowPositionTests(unittest.TestCase):
    def test_ignores_saved_position_preserves_size(self):
        geometry, minimum = startup_geometry('1200x850+50+60', (1920, 0, 1920, 1080))
        self.assertEqual(geometry, '1200x850+2280+100')
        self.assertEqual(minimum, (1050, 820))

    def test_left_and_above_displays(self):
        self.assertEqual(startup_geometry('1180x900', (-1920, -1080, 1920, 1080))[0],
                         '1180x900+-1550+-1005')

    def test_small_display_fits_window(self):
        geometry, minimum = startup_geometry('1180x900', (0, 0, 1024, 768))
        self.assertEqual(geometry, '984x668+20+35')
        self.assertEqual(minimum, (984, 668))

    def test_invalid_saved_geometry(self):
        self.assertEqual(startup_geometry('invalid', (0, 0, 1920, 1080))[0],
                         '1180x900+370+75')

    def test_native_lookup_failure_falls_back(self):
        root = Mock()
        root.winfo_screenwidth.return_value = 1920
        root.winfo_screenheight.return_value = 1080
        with patch('court_reserv.ui.window_position.sys.platform', 'darwin'), patch(
            'court_reserv.ui.window_position.pointer_display_bounds', side_effect=RuntimeError
        ):
            position_startup_window(root, '1180x900+2000+0')
        root.geometry.assert_called_once_with('1180x900+370+75')
