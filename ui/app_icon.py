import os
import sys

from PySide6.QtGui import QIcon


def get_app_icon() -> QIcon:
    """Return application icon for both dev and PyInstaller environments."""
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", "")
        icon_name = "lioil.ico" if os.name == "nt" else "lioil.icns"
        icon_path = os.path.join(base_path, icon_name)
        if os.path.exists(icon_path):
            return QIcon(icon_path)

    icon_name = "lioil.ico" if os.name == "nt" else "lioil.icns"
    if os.path.exists(icon_name):
        return QIcon(icon_name)

    return QIcon()
