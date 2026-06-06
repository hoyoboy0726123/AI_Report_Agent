"""全域快捷鍵監聽管理。"""

import keyboard


class HotkeyManager:
    """以註冊/解除模式管理單一快捷鍵；不需自行 polling。"""

    def __init__(self, hotkey, callback):
        self.hotkey = hotkey
        self.callback = callback
        self._handle = None

    def start(self):
        if self._handle is None:
            self._handle = keyboard.add_hotkey(self.hotkey, self.callback)

    def stop(self):
        if self._handle is not None:
            try:
                keyboard.remove_hotkey(self._handle)
            except (KeyError, ValueError):
                pass
            self._handle = None
