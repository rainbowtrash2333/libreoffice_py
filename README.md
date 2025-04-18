# libreoffice_py

My job is statistics, I have to deal with a lot of boring tables and reports. For particular reasons, I am limited to using a Linux computer, not a Windows one, I can't use Office VBA for this work. SO I created this Simple Python tool for manipulating LibreOffice components.

## Project Installation

[poetry](https://python-poetry.org/) is used to install this project.If you have not used poetry before, read [this](https://python-poetry.org/docs/basic-usage/).

Install poetry.
```sh
pip install poetry
```

Clone this project.

```sh
git clone https://github.com/rainbowtrash2333/libreoffice_py.git
cd libreoffice_py
```

### Windows
install LibreOffice
```sh
py -3.9 -m venv .\.venv
.\.venv\Scripts\activate
poetry install
```
switch to  UNO evnvironment.
```sh
oooenv env -t
oooenv env -u
# UNO Environment
```
if you need use poetry, switch back
```sh
oooenv env -t
poetry <some-command>
oooenv env -t
```
## Linux

```sh
python3 -m venv .venv
source .venv/bin/activate
poetry install
```

### Testing the Installed Environment

If there is no import error, then you should be good to go.

```sh
python3 # if windows, run `python`
>>> import uno
>>> exit()
```
 