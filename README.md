<h3 align="center">Excel data retriever</h3>

<div align="center">

[![Status](https://img.shields.io/badge/status-active-success.svg)]()
[![GitHub Issues](https://img.shields.io/github/issues/ForsakenWing/ExcelToRawData.svg)](https://github.com/ForsakenWing/ExcelDataRetriever/issues)
[![GitHub Pull Requests](https://img.shields.io/github/issues-pr/ForsakenWing/ExcelToRawData.svg)](https://github.com/ForsakenWing/ExcelDataRetriever/pulls)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](/LICENSE)

</div>

---

<p align="center"> Use this <b>code</b> to parse data from your xml files and reuse it for your purposes. In <b>templates</b> folder you can see how to use it for automation purposes.
    <br> 
</p>

## 📝 Table of Contents

- [About](#about)
- [Getting Started](#getting_started)
- [Deployment](#deployment)
- [Usage](#usage)
- [Built Using](#built_using)
- [TODO](../TODO.md)
- [Contributing](../CONTRIBUTING.md)
- [Authors](#authors)
- [Acknowledgments](#acknowledgement)

## 🧐 About <a name = "about"></a>

Created this project because I like challenges. Big appreciation to my mentor [Ay](https://github.com/Umutayb).

## 🏁 Getting Started <a name = "getting_started"></a>

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See [deployment](#deployment) for notes on how to deploy the project on a live system.

### Prerequisites

<b>Python 3.10+
Git 2.0+
</b>

<i>Ubuntu-based systems</i>
```
sudo apt install python3.10
```
<i>MacOS systems</i>
```
brew install python@3.10
```
<i>Windows systems</i>
```
python-3.10.0.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
```
### Installing

A step by step series of examples that tell you how to get a development env running.

Clone repository and change directory

```
git clone https://github.com/ForsakenWing/ExcelDataRetriever
cd ExcelDataRetriever/
```

Set up virtual environment and install requirements

```
python3.10 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Run application as CLI

```
CLI=1 python main.py -h
```

Run application as web-service without docker
```
gunicorn --bind 0.0.0.0:8091 -k uvicorn.workers.UvicornWorker main:app
```
Run application as web-service with docker
```
docker build -t excel_parser .
docker run -d -p 8075:8000 excel_parser
```

## 🔧 Running the tests <a name = "tests"></a>

Explain how to run the automated tests for this system.

### Break down into end to end tests

Explain what these tests test and why

```
Give an example
```

### And coding style tests

Explain what these tests test and why

```
Give an example
```

## 🎈 Usage <a name="usage"></a>

Add notes about how to use the system.

## 🚀 Deployment <a name = "deployment"></a>

Add additional notes about how to deploy this on a live system.

## ⛏️ Built Using <a name = "built_using"></a>

- [MongoDB](https://www.mongodb.com/) - Database
- [Express](https://expressjs.com/) - Server Framework
- [VueJs](https://vuejs.org/) - Web Framework
- [NodeJs](https://nodejs.org/en/) - Server Environment

## ✍️ Authors <a name = "authors"></a>

- [@ForsakenWing](https://github.com/ForsakenWing) - Initial work
- [@Umutayb](https://github.com/Umutayb) - Idea

See also the list of [contributors](https://github.com/ForsakenWing/ExcelDataRetriever/contributors) who participated in this project.

## 🎉 Acknowledgements <a name = "acknowledgement"></a>

- Inspiration
- References
