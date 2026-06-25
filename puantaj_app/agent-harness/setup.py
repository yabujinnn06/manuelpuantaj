from setuptools import setup, find_namespace_packages

setup(
    name="cli-anything-puantaj",
    version="1.0.0",
    description="Agent-usable CLI for the Rainstaff puantaj application",
    packages=find_namespace_packages(include=["cli_anything.*"]),
    python_requires=">=3.10",
    install_requires=["click>=8.0"],
    extras_require={"report": ["openpyxl>=3.1", "Pillow>=11.0", "reportlab>=4.0"]},
    entry_points={
        "console_scripts": [
            "cli-anything-puantaj=cli_anything.puantaj.puantaj_cli:main",
        ],
    },
)
