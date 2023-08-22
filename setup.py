from cx_Freeze import setup, Executable
executables = [Executable('main.py', base='Win32GUI',
                          target_name='AccountCheck.exe',
                          icon='ico/analysis_finance_statistics_business_graph_chart_report_icon_254045.ico')]
excludes = ['unittest', 'asyncio', 'sqlite3', 'distutils', 'concurrent']
includefiles = ['__VERS__.txt', 'images']

inc_paclages = ['PIL', 'PySimpleGUI','xlsxwriter','certifi','charset_normalizer','clr','cx-Freeze','cx-Logging','et_xmlfile','idna',
 'lief', 'lxml', 'microsoft', 'numpy', 'openpyxl', 'pandas', 'pip', 'psutil', 'python_dateutil', 'pytils',
 'pytz', 'pywin32', 'requests', 'setuptools', 'six', 'urllib3', 'settings', 'collections', 'ctypes', 'curses', 'dateutil', 'encodings',
                'html', 'http', 'importlib', 'json', 'lib2to3', 'logging', 'multiprocessing', 'pkg_resources', 'pydoc_data',
                'pywin', 'pywin32_system32', 're', 'tcl8', 'tcl8.6', 'test', 'tk8.6', 'tkinter', 'unittest', 'urllib', 'win32com', 'xml', 'xmlrpc', 'email']


options = {
      'build_exe': {
          'include_files': includefiles,
          'includes': ['pptx'],
            'excludes': excludes,
            'build_exe': 'build_windows',
            # 'zip_include_packages': zip_include_packages,
            "zip_include_packages": inc_paclages,
            # "zip_exclude_packages": "",
            'optimize': 1
      }
}


setup(name='AccountCheck',
      version='1.0.1',
      description='Сверка',
      executables=executables,
      options=options)