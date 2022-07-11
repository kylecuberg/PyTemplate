from setuptools import find_packages
from setuptools import setup


with open("README.md") as f:
    readme = f.read()

with open("LICENSE") as f:
    license = f.read()

setup(
    name="Basic",
    version="0.1.1",
    description="Basic package",
    long_description=readme,
    author="Kyle Patterson",
    url="https://github.com/kylecuberg",
    license=license,
    packages=find_packages(exclude=("Tests", "Docs", "Results")),
)
