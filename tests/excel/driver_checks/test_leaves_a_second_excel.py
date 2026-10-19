"""A check of the driver: an Excel a test starts and reports, as conftest.py's direct_excel
does, that is still running once pytest has exited must be waited for, ended, and the run
failed."""

from conftest import report_excel


def test_leaves_a_second_excel(run_folder):
    import win32process
    from win32com.client import DispatchEx

    app = DispatchEx("Excel.Application")
    report_excel(run_folder, "left", win32process.GetWindowThreadProcessId(app.Hwnd)[1])
    app.UserControl = True  # It then outlives its release: an Excel that never quits.
