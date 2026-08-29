# Detection of date/time number formats and conversion of Excel serial date
# values. A cell value in an xls file is just a Float64; whether it denotes a
# date, a time or both is determined by the number format of the cell's XF
# record.

"""
    CellFormatKind

Classification of a cell's number format: `FORMAT_NONE` for ordinary numbers,
`FORMAT_DATE` for date-only formats, `FORMAT_TIME` for time-only formats and
`FORMAT_DATETIME` for formats with both date and time components.
"""
@enum CellFormatKind begin
    FORMAT_NONE
    FORMAT_DATE
    FORMAT_TIME
    FORMAT_DATETIME
end

# Builtin number format ids that denote dates or times (same set xlrd uses):
# 14-17 dates, 18-21 times, 22 date+time, 27-36 East Asian dates, 45-47
# elapsed times, 50-58 East Asian date variants.
const BUILTIN_FORMAT_KINDS = Dict{Int,CellFormatKind}(
    Dict(i => FORMAT_DATE for i in [14:17; 27:36; 50:58])...,
    Dict(i => FORMAT_TIME for i in [18:21; 45:47])...,
    22 => FORMAT_DATETIME,
)

# Classify a custom number format string. Quoted literals, escaped characters,
# padding/fill markers and bracket sections (colors like [Red], conditions
# like [<=100]) must be ignored; elapsed-time tokens like [h] or [ss] count as
# time. In what remains, y and d denote a date component, h and s a time
# component, and m either of the two: minutes when next to an h or s code,
# months otherwise.
function format_string_kind(fmt::AbstractString)
    stripped = IOBuffer()
    i = firstindex(fmt)
    n = lastindex(fmt)
    while i <= n
        c = fmt[i]
        if c == '"'
            j = findnext(isequal('"'), fmt, nextind(fmt, i))
            j === nothing && break
            i = nextind(fmt, j)
        elseif c == '\\' || c == '_' || c == '*'
            i = nextind(fmt, i)
            i <= n && (i = nextind(fmt, i))
        elseif c == '['
            j = findnext(isequal(']'), fmt, i)
            j === nothing && break
            inner = SubString(fmt, nextind(fmt, i), prevind(fmt, j))
            occursin(r"^[hms]+$"i, inner) && print(stripped, inner)
            i = nextind(fmt, j)
        else
            print(stripped, c)
            i = nextind(fmt, i)
        end
    end
    codes = String(take!(stripped))

    # AM/PM markers denote a time component but their m is not a month.
    hastime = occursin(r"AM/PM|A/P"i, codes)
    codes = replace(codes, r"AM/PM|A/P"i => "")

    hastime |= occursin(r"[hs]"i, codes)
    hasdate = occursin(r"[yd]"i, codes)
    # A run of m codes not adjacent to an h or s code means months, not
    # minutes. The lookarounds exclude m itself so that only complete runs
    # are considered.
    hasdate |= occursin(r"(?<![hsm])m+(?![hsm])"i, replace(codes, r"[^a-zA-Z]" => ""))

    hasdate && hastime && return FORMAT_DATETIME
    hasdate && return FORMAT_DATE
    hastime && return FORMAT_TIME
    return FORMAT_NONE
end

function format_kind(index::Integer, custom_formats::Dict{UInt16,String})
    # A FORMAT record can redefine any index, including builtin ones, so the
    # formats stored in the file take precedence over the builtin table.
    haskey(custom_formats, index) && return format_string_kind(custom_formats[index])
    return get(BUILTIN_FORMAT_KINDS, Int(index), FORMAT_NONE)
end

const MILLISECONDS_PER_DAY = 86_400_000

"""
    excel_serial_to_temporal(value, kind, is1904)

Convert an Excel serial date/time value to a `Date`, `DateTime` or `Time`,
depending on the format `kind` of the cell and the value: a date-only format
with no fractional part gives a `Date`, a time-only format with a value below
one day gives a `Time`, and everything else gives a `DateTime`. In the 1900
date system the nonexistent date 1900-02-29 (serial 60) maps to 1900-02-28.
Negative values are not valid dates and are returned unchanged as `Float64`.
"""
function excel_serial_to_temporal(value::Float64, kind::CellFormatKind, is1904::Bool)
    value < 0 && return value
    days = floor(Int, value)
    ms = round(Int, (value - days) * MILLISECONDS_PER_DAY)
    if ms >= MILLISECONDS_PER_DAY
        days += 1
        ms = 0
    end
    t = Time(0) + Millisecond(ms)
    days == 0 && kind != FORMAT_DATE && return t
    if is1904
        d = Date(1904, 1, 1) + Day(days)
    elseif days == 60
        d = Date(1900, 2, 28)
    elseif days < 60
        d = Date(1899, 12, 31) + Day(days)
    else
        d = Date(1899, 12, 30) + Day(days)
    end
    kind == FORMAT_DATE && ms == 0 && return d
    return DateTime(d, t)
end
