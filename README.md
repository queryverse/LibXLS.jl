# LibXLS

[![Project Status: Active - The project has reached a stable, usable state and is being actively developed.](http://www.repostatus.org/badges/latest/active.svg)](http://www.repostatus.org/#active)
[![Build Status](https://github.com/queryverse/LibXLS.jl/actions/workflows/juliaci.yml/badge.svg?branch=main)](https://github.com/queryverse/LibXLS.jl/actions/workflows/juliaci.yml)

## Overview

LibXLS reads legacy Excel xls files (the BIFF format used by Excel 97-2003).
It wraps the [libxls](https://github.com/libxls/libxls) C library, with
binaries provided by `libxls_jll`, so it works without any Python or Java
dependency.

For modern xlsx files use [XLSX.jl](https://github.com/JuliaData/XLSX.jl)
instead; LibXLS deliberately only handles the legacy format. If you want one
package that reads both formats behind a single API, use
[ExcelReaders.jl](https://github.com/queryverse/ExcelReaders.jl), which is
built on LibXLS and XLSX.jl.

## Installation

```julia
Pkg.add("LibXLS")
```

## Getting started

```julia
using LibXLS

wb = openxls("data.xls")

sheetnames(wb)                   # names of all sheets
ws = getworksheet(wb, "Sheet1")  # or by index: getworksheet(wb, 1)

nrows, ncols = size(ws)
ws[1, 1]                         # value of the cell in the first row and column

close(wb)
```

The do-block form closes the workbook automatically:

```julia
openxls("data.xls") do wb
    getworksheet(wb, 1)[2, 3]
end
```

## API

### Workbooks

* `openxls(filepath)` opens an xls file and returns a `Workbook`. The file
  format is validated from the file's content; opening an xlsx file gives an
  error pointing to XLSX.jl. `openxls(f, filepath)` calls `f` on the workbook
  and closes it afterwards.
* `close(wb)` releases the resources held by the C library; `isopen(wb)`
  reports whether the workbook is still open. Workbooks also close themselves
  when garbage collected.
* `sheetcount(wb)` — number of sheets, including hidden and empty ones.
* `sheetnames(wb)` — names of all sheets, in sheet order.
* `LibXLS.sheetname(wb, i)` / `LibXLS.sheetindex(wb, name)` — translate
  between sheet indices and names.
* `LibXLS.isvisible(wb, index_or_name)` — whether a sheet is visible.
* `LibXLS.is1904(wb)` — whether the file uses the 1904 date system (cell
  values already account for this).

### Worksheets

* `getworksheet(wb, index_or_name)` returns a `Worksheet`; `wb[1]` and
  `wb["Sheet1"]` are shorthands. Worksheets are parsed on first access and
  cached.
* `size(ws)` — dimensions of the used cell range as `(rows, columns)`.
* `ws[row, col]` — the value of a cell, using 1-based indices.
* `LibXLS.sheetname(ws)` / `LibXLS.sheetindex(ws)` — the sheet's name/index.

### Cell values

`ws[row, col]` returns plain Julia values:

| Excel cell                              | Julia value          |
| --------------------------------------- | -------------------- |
| blank                                   | `missing`            |
| number                                  | `Float64`            |
| text                                    | `String`             |
| boolean                                 | `Bool`               |
| date (date-only number format)          | `Dates.Date`         |
| date + time                             | `Dates.DateTime`     |
| time (time-only format, less than 24h)  | `Dates.Time`         |
| error (`#DIV/0!`, `#N/A`, ...)          | `CellError`          |

Formula cells return the cached result of the formula, mapped by the same
rules.

An xls file stores dates and times as plain numbers; whether a number denotes
a date is determined by the cell's number format, both for the builtin
formats and for custom format strings. Serial values convert correctly in
both the 1900 date system (including Excel's phantom 1900-02-29, which maps
to 1900-02-28) and the 1904 date system.

`CellError` wraps the BIFF error code of a cell holding an Excel error value
and prints as the corresponding error string:

| Code   | Error     |
| ------ | --------- |
| `0x00` | `#NULL!`  |
| `0x07` | `#DIV/0!` |
| `0x0F` | `#VALUE!` |
| `0x17` | `#REF!`   |
| `0x1D` | `#NAME?`  |
| `0x24` | `#NUM!`   |
| `0x2A` | `#N/A`    |

## Limitations

* Reading only — for writing spreadsheets use
  [XLSX.jl](https://github.com/JuliaData/XLSX.jl) (xlsx).
* Legacy xls only; xlsx files are rejected with a pointer to XLSX.jl.
* Defined names are not exposed by the underlying C library.
* Password-protected/encrypted workbooks are not supported.
