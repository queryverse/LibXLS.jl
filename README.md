# LibXLS

[![Project Status: Active - The project has reached a stable, usable state and is being actively developed.](http://www.repostatus.org/badges/latest/active.svg)](http://www.repostatus.org/#active)
[![Build Status](https://github.com/queryverse/LibXLS.jl/actions/workflows/juliaci.yml/badge.svg?branch=main)](https://github.com/queryverse/LibXLS.jl/actions/workflows/juliaci.yml)

## Overview

LibXLS reads legacy Excel xls files (the BIFF format used by Excel 97-2003).
It wraps the [libxls](https://github.com/libxls/libxls) C library, with
binaries provided by `libxls_jll`, so it works without any Python or Java
dependency.

For modern xlsx files use [XLSX.jl](https://github.com/JuliaData/XLSX.jl)
instead; LibXLS deliberately only handles the legacy format.

## Usage

```julia
using LibXLS

wb = openxls("data.xls")

sheetnames(wb)          # names of all sheets
ws = getworksheet(wb, "Sheet1")  # or by index: getworksheet(wb, 1)

nrows, ncols = size(ws)
ws[1, 1]                # value of the cell in the first row and column

close(wb)
```

The do-block form closes the workbook automatically:

```julia
openxls("data.xls") do wb
    getworksheet(wb, 1)[2, 3]
end
```

Cell values are returned as `Float64`, `String`, `Bool`, `DateTime`, `Time`,
`CellError` (for cells holding an Excel error such as `#DIV/0!`) or `missing`
(for blank cells). Whether a numeric cell holds a date is determined from the
cell's number format, honoring both the 1900 and the 1904 date system.
