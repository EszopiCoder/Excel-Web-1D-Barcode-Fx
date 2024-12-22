# Excel Web 1D Barcode Functions
## Purpose and Features
Custom function library to generate the following 1D barcodes using `=SPARKLINE()`:
| Barcode Type | Barcodes | Status |
| --- | --- | --- |
| 1D Code | Code 11, Code 39, Code 93, ~~Code 128~~ | In progress |
| 1D UPC/EAN | EAN-2, EAN-5, EAN-8, EAN-13, UPC-A, UPC-E | Completed |
| 1D ITF | ITF, ITF-14 | Completed |
## Installation and Usage
- Install the add-in by downloading [manifest.xml](https://github.com/EszopiCoder/Excel-Web-1D-Barcode-Fx/blob/main/manifest.xml) and upload it. See detailed instructions [here](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#manually-sideload-an-add-in-to-office-on-the-web)
- All check digits are auto-calculated if not included and verified if included.

| Function Name | Data Types | Length and Format | Check Digit |
| --- | --- | --- | --- |
| `Code11()` | Numeric and dash | Unlimited | None |
| `Code39()` | Uppercase alphanumeric, space, and -$%./+ | Unlimited + check digit (optional) | Modulo 43 |
| `Code93()` | Uppercase alphanumeric, space, and -$%./+ | Unlimited + 2 check digits | Modulo 47 |
| `EAN_2()` | Numeric | 2 digits | None |
| `EAN_5()` | Numeric | 5 digits | None |
| `EAN_13()` | Numeric | 12 digits + check digit | GS1 check digit |
| `ITF()` | Numeric | Unlimited | None |
| `ITF_14()` | Numeric | 13 digits + check digit | GS1 check digit |
| `UPCA()` | Numeric | 11 + check digit (UPC-A) or 8 digits (EAN-8) | GS1 check digit |
| `UPCE()` | Numeric | ("0" or "1") + 6 digits | None |
