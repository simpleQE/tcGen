<h1 align="center"> TCGEN</h1>
<h3 align="center">Test Case Generation Tool</h3>

<p align="center">
<a href="https://pypi.org/project/tcgen"><img alt="Python Versions"
src="https://img.shields.io/pypi/pyversions/miqsel.svg?style=flat"></a>
<a href="https://github.com/SurajJadhav7/tcGen/blob/master/LICENSE"><img alt="License: GPLV3"
src="https://img.shields.io/pypi/l/miqsel.svg?version=latest"></a>
<a href="https://pypi.org/project/black"><img alt="Code style: black"
src="https://img.shields.io/badge/code%20style-black-000000.svg"></a>
</p>

Simple command line tool to generate manual test cases.


### Installation:

1. Clone the project repository
```bash
git clone https://github.com/SurajJadhav7/tcGen.git
```
2. Install project using below command

```bash
python setup.py install
```


## Usage:

```shell
>>> tcgen
Usage: tcgen [OPTIONS] COMMAND [ARGS]...

  Test case generator tool

Options:
  --version  Show the version and exit.
  --help     Show this message and exit.

Commands:
  generate  Provide url

>>> tcgen generate --help
Usage: tcgen generate [OPTIONS]

  Provide url

Options:
  -u, --url TEXT  URL to generate test cases
  --help          Show this message and exit.

```

*Generate manual test cases*

```bash
tcgen generate -u <url>
```

**NOTE: User will find test cases sheet in file -/output.xlsx**
