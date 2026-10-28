"""Read-only checks; does not start Origin/PowerPoint or inspect user documents."""
import importlib
import json
import platform
import sys


def check():
    result = {'windows': sys.platform == 'win32', 'python': platform.python_version(),
              'python_bits': platform.architecture()[0], 'modules': {}, 'com_registered': {}}
    for name in ('tkinter', 'pythoncom', 'win32com.client', 'lxml', 'PIL', 'numpy'):
        try:
            importlib.import_module(name)
            result['modules'][name] = True
        except ImportError:
            result['modules'][name] = False
    if result['windows']:
        import winreg
        for name in ('Origin.Application', 'PowerPoint.Application'):
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, name + '\\CLSID'):
                    result['com_registered'][name] = True
            except OSError:
                result['com_registered'][name] = False
    result['note'] = 'Registration only; licensing, activation, application versions and native export require manual validation.'
    return result


if __name__ == '__main__':
    print(json.dumps(check(), indent=2))
