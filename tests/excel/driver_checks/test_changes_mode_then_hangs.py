"""A check of the driver: a run that changes Windows' app mode and is then stopped with
Ctrl+C, or reaches its timeout, must still leave the app mode as it was."""

PERSONALIZE = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"


def test_changes_mode_then_hangs(book):
    import winreg  # Here, not at the top: off Windows the module must still be collected.

    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, PERSONALIZE, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
        try:
            light = winreg.QueryValueEx(key, "AppsUseLightTheme")[0]
        except FileNotFoundError:
            light = 1
        winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, 0 if light else 1)
    book.script("hang.py", "import time\n\ndef hang():\n    while True:\n        time.sleep(1)\n")
    book.formula("A1", '=PyCall("hang.py", "hang")')
