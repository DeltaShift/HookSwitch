# HookSwitch

High-performance, low-memory XLSX ⇄ CSV conversion engine written in pure PHP using streaming architecture.

HookSwitch is designed for production environments where large spreadsheet files must be converted safely without exhausting memory. It implements a streaming parser and writer similar to modern SaaS systems, without relying on external libraries.

Owner: https://github.com/mizukaze554

---

# Features

* Convert XLSX → CSV
* Convert CSV → XLSX
* Pure PHP (no frameworks, no external dependencies)
* Streaming processing (constant low memory usage)
* Handles very large files (100MB, 500MB, 1GB+)
* Disk-based shared string indexing
* XML-safe writing and sanitization
* Automatic CSV delimiter detection
* Performance benchmarking support
* Secure path validation

---

# Architecture

HookSwitch uses a streaming model:

XLSX → CSV flow:

* Open XLSX as ZIP archive
* Stream worksheet XML using XMLReader
* Read row-by-row
* Write directly to CSV

CSV → XLSX flow:

* Read CSV row-by-row
* Stream XML worksheet creation
* Package into XLSX ZIP container

Memory usage stays nearly constant regardless of file size.

---

# Requirements

PHP 8.0 or higher

Required PHP extensions:

* zip
* xmlreader

Install on Ubuntu:

```bash
sudo apt install php-zip php-xml
```

---

# Project Structure

```
hookswitch/

file_format_change.php
test_convert.php
test.xlsx
test.csv
README.md
```

---

# Usage

Run conversion from terminal:

CSV → XLSX

```bash
php test_convert.php input.csv output.xlsx csv_to_xlsx
```

XLSX → CSV

```bash
php test_convert.php input.xlsx output.csv xlsx_to_csv
```

Example:

```bash
php test_convert.php test.xlsx result.csv xlsx_to_csv
```

---

# Example Output

```
SUCCESS: conversion completed

Wall time (s): 2.314582
CPU usage (%): 97.42
Peak memory (MB): 8.50
Output size (bytes): 5242880
```

---

# Performance

Typical performance on modern server:

| File Size | Memory  | Time      |
| --------- | ------- | --------- |
| 10MB      | 5-10MB  | < 1 sec   |
| 100MB     | 5-15MB  | 10-20 sec |
| 500MB     | 10-20MB | 1-2 min   |

Memory does not grow with file size.

---

# Security

HookSwitch includes:

* Path traversal protection
* XML sanitization
* Invalid character filtering
* Safe filesystem access restrictions

---

# Integration Example

```php
require_once 'file_format_change.php';

FileFormatChangeService::convertXlsxToCsv(
    'input.xlsx',
    'output.csv'
);
```

---

# Why HookSwitch Exists

Most PHP spreadsheet libraries load entire files into memory, which crashes servers at scale.

HookSwitch uses streaming to provide:

* predictable performance
* low memory usage
* scalability

Suitable for SaaS, APIs, and high-load backend systems.

---

# Roadmap

Future improvements:

* Parallel worker support
* Queue integration
* Multi-sheet support
* Progress tracking
* HTTP API service

---

# License

MIT License

---

# Author

GitHub:
https://github.com/mizukaze554

Project:
HookSwitch

---

# Contributing

Pull requests are welcome.

---

# Star the project

If HookSwitch helps you, please star the repository.
