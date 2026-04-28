"""GUI 兼容入口。

实际 App 实现位于 .gui_app，后续子模块位于 .gui_modules。
"""

from .gui_app import App, run_gui

__all__ = ['App', 'run_gui']
