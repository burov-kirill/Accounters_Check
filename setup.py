from cx_Freeze import setup, Executable
executables = [Executable('main.py', base='Win32GUI',
                          target_name='AccountCheck.exe',
                          icon='ico/analysis_finance_statistics_business_graph_chart_report_icon_254045.ico')]
excludes = ['unittest', 'asyncio', 'sqlite3', 'distutils']
includefiles = ['__VERS__.txt', 'images']




options = {
      'build_exe': {
          'include_files': includefiles,
            'excludes': excludes,
            'build_exe': 'build_windows',
            # 'zip_include_packages': zip_include_packages,
            "zip_include_packages": "*",
            "zip_exclude_packages": "",
            'optimize': 1
      }
}


setup(name='AccountCheck',
      version='1.0.1',
      description='Сверка',
      executables=executables,
      options=options)