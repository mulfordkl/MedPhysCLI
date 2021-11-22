from setuptools import setup

setup(
    name='medphys-cli',
    version='0.1',
    py_modules=['create_report'],
    install_requires=[
        'Click',
        'openpyxl'
    ],
    entry_points='''
        [console_scripts]
        create_report=create_report:cli
    ''',
)