from setuptools import setup

setup(
    name="outlook-backup",
    version="0.0.1",
    url="https://github.com/andiwand/outlook-backup",
    author="Andreas Stefl",
    install_requires=[],
    author_email="stefl.andreas@gmail.com",
    description="Backup and restore script for Microsoft Outlook settings.",
    long_description="",
    package_dir={"": "src"},
    packages=["outlookbackup"],
    platforms=["windows"],
    entry_points={
        "console_scripts": ["outlook-backup = outlookbackup.cli:main"]
    },
)
