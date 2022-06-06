from fbs_runtime.application_context.PySide2 import ApplicationContext

import sys

from package.main_window import MainWindow

if __name__ == '__main__':
    app = ApplicationContext()
    window = MainWindow(ctx=app)
    window.resize(500, 500)
    window.show()
    exit_code = app.app.exec_()
    sys.exit(exit_code)
