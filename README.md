# Fastest Way to Read Excel in Python

Compare ways to read Excel files in Python.

The repo includes the source files for running the benchmarks presented in the article ["Fastest Way to Read Excel in Python"](https://hakibenita.com/fast-excel-python).

## Setup

Create a virtual environment and install dependencies:

```bash
$ python -m venv venv
$ source venv/bin/activate
(venv) $ pip install -r requirements.txt
```

## Running the Benchmark

To run the benchmark execute the following command:

```bash
(venv) $ python benchmark.py
```

The repo includes two Excel files:

- `file.xlsx`: large file used for the benchmark

- `file-sample.xlsx`: smaller file to use for development
