from setuptools import setup

setup(
    name="aimath2docx",
    version="1.0",
    py_modules=["aimath2docx"],
    install_requires=[
        "python-docx",
        "latex2mathml",
        "lxml",
    ],
    entry_points={
        "console_scripts": [
            "aimath2docx = aimath2docx:main"
        ]
    }
)
