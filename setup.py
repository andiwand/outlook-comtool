from setuptools import setup

setup(
    name="outlook-comtool",
    version="0.0.1",
    url="https://github.com/andiwand/outlook-comtool",
    author="Andreas Stefl",
    install_requires=[
        'win32com;platform_system=="Windows"'
    ],
    author_email="stefl.andreas@gmail.com",
    description="Microsoft Outlook win32com script collection.",
    long_description="",
    package_dir={"": "src"},
    packages=["outlookcomtool"],
    platforms=["windows"],
    entry_points={
        "console_scripts": ["outlook-dumpcontacts = outlookbackup.dumpcontactscli:main"],
		"console_scripts": ["outlook-filtercontacts = outlookbackup.filtercontactscli:main"]
    },
)
