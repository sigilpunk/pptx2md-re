import setuptools
import re
from pathlib import Path

FALLBACK_VERSION = "0.0.1"

with open("./pptx2md_re/main.py", "r") as f:
    data = f.read()
    version = re.findall(r'PPTX2MD_VERSION\s=\s\"(.+?)\"', data)
    version = version[0] if version[0] else FALLBACK_VERSION

setuptools.setup(
    name="pptx2md-re",
    version=version,
    author="sigilpunk",
    author_email="jskresl@gmail.com",
    packages=["pptx2md_re"],
    description="Suite of tools to convert PowerPoint Presentstions (.pptx) to other formats (.json, .md, .pptxt)",
    url="https://github.com/sigilpunk/pptx2md",
    license='MIT',
    python_requires='>=3.8',
    install_requires=['pathlib', 'python-pptx', 'tqdm']
)