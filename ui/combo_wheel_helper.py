from __future__ import annotations

from PySide6.QtCore import QObject, QEvent
from PySide6.QtWidgets import QComboBox


class ComboWheelController(QObject):
    """讓一般 QComboBox 在滾輪時切換清單項目，而非誤改輸入欄位數字。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._targets: dict[object, QComboBox] = {}

    @staticmethod
    def _steps_from_wheel(event) -> int:
        delta = event.angleDelta().y()
        if delta == 0:
            return 0
        steps = int(delta / 120)
        if steps == 0:
            steps = 1 if delta > 0 else -1
        return steps

    def register_combo(self, combo: QComboBox):
        # 自訂子類通常已自行處理滾輪，避免覆蓋既有行為。
        if type(combo) is not QComboBox:
            return

        objects = [combo]
        line_edit = combo.lineEdit()
        if line_edit is not None:
            objects.append(line_edit)

        view = combo.view()
        if view is not None:
            objects.append(view)
            viewport = view.viewport()
            if viewport is not None:
                objects.append(viewport)

        for obj in objects:
            if obj not in self._targets:
                self._targets[obj] = combo
                obj.installEventFilter(self)

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Wheel:
            combo = self._targets.get(watched)
            if combo is not None and combo.isEnabled() and combo.count() > 0:
                steps = self._steps_from_wheel(event)
                if steps != 0:
                    current_index = combo.currentIndex()
                    if current_index < 0:
                        current_index = 0
                    target_index = current_index - steps
                    if target_index < 0:
                        target_index = 0
                    elif target_index >= combo.count():
                        target_index = combo.count() - 1
                    if target_index != combo.currentIndex():
                        combo.setCurrentIndex(target_index)
                event.accept()
                return True

        return super().eventFilter(watched, event)


def attach_combo_wheel_behavior(root) -> ComboWheelController:
    controller = getattr(root, "_combo_wheel_controller", None)
    if controller is None:
        controller = ComboWheelController(root)
        setattr(root, "_combo_wheel_controller", controller)

    for combo in root.findChildren(QComboBox):
        controller.register_combo(combo)

    return controller
