# LibXLS.jl v1.0.0 Release Notes

* Complete rewrite as a native reader for legacy Excel xls files, wrapping
  the libxls C library via the registered `libxls_jll` binaries (no
  BinaryProvider, no build step).
* Full cell value support: numbers, strings, booleans, blank cells
  (`missing`), error cells (new `CellError` type) and cached formula results.
* Date and time support: cell number formats (builtin ids and custom format
  strings) determine whether a number is a `Date`, `DateTime` or `Time`,
  honoring both the 1900 and the 1904 date system.
* Worksheet access via `getworksheet`/`wb[...]`, `size` and `ws[row, col]`.
* Tests use the test item framework; minimum supported Julia is 1.12.

# LibXLS.jl v0.0.1 Release Notes
* Initial release
