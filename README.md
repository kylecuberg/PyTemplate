# PyTemplate
Template Python Project

## Table of Contents
- [About](#about)
- [Requirements](#requirements)
- [Usage](#usage)
- [TODO](#todo)
- [Authors](#authors)
- [License](#license)
- [CodeStyle](#codestyle)
- [Badges](#badges)

## About
General Python template project.

## Requirements

  * Python 3.7 or greater

> This code makes use of the `f"..."` or [f-string syntax](https://www.python.org/dev/peps/pep-0498/). This syntax was introduced in Python 3.6.

For additional requirements and dependencies, please see the requirements.txt file.

## Deployment / Usage
To create virtual environment:
> py -m venv env

To activate
> env\Scripts\activate

To install requirements.txt
> pip install -r requirements.txt

To freeze requirements.txt after additional installs
> pip freeze > requirements.txt

To apply black (depends on system setup)
> py -m black [filename]

> black [filename]

To analyze flake8 violations
> flake8 Project --output-file ./reports/flake8/flake8stats.txt

> flake8 Project --exit-zero --format=html --htmldir ./reports/flake8 --statistics --tee --output-file flake8stats.txt

For genbadge on Flake8 (must do both flake8 commands first)
> genbadge flake8 --output-file ./reports/flake8/badge.svg


## TODO
- [] Item 1
- [] Item 2
- [] Item 3

## Authors
- [Kyle Patterson](https://github.com/kylekap)


### License
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](/LICENSE)

### CodeStyle
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
[![Flake8 Status](./reports/flake8/badge.svg)](./reports/flake8/index.html)
### Badges
[![Windows](https://svgshare.com/i/ZhY.svg)](https://svgshare.com/i/ZhY.svg)
[![Status](https://img.shields.io/badge/status-active-success.svg)]()
[![GitHub Issues](https://img.shields.io/github/issues/kylecuberg/PyTemplate.svg)](https://github.com/kylecuberg/PyTemplate/issues)
[![GitHub Pull Requests](https://img.shields.io/github/issues-pr/kylecuberg/PyTempalte.svg)](https://github.com/kylecuberg/PyTemplate/pulls)
[![Python 3.7](https://img.shields.io/badge/python-3.7-blue.svg)](https://www.python.org/downloads/release/python-370/)
[![Python 3.8](https://img.shields.io/badge/python-3.8-blue.svg)](https://www.python.org/downloads/release/python-380/)
[![Python 3.9](https://img.shields.io/badge/python-3.9-blue.svg)](https://www.python.org/downloads/release/python-390/)
[![Python 3.10](https://img.shields.io/badge/python-3.10-blue.svg)](https://www.python.org/downloads/release/python-3100/)

