from setuptools import setup

setup(
    name="aimath2docx",
    version="1.0",
    py_modules=["aimath2docx"],
    install_requires=[
        "python-docx",
        "latex2mathml",
        "lxml",
        "customtkinter",
        "mathml2omml_as"
    ],
    entry_points={
        "console_scripts": [
            "aimath2docx = aimath2docx:main",
            "aimath2docx-gui = aimath2docx-gui:main"
        ]
    }
)
