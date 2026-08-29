# Detection of date/time number formats and conversion of Excel serial date
# values. A cell value in an xls file is just a Float64; whether it denotes a
# date or time is determined by the number format of the cell's XF record.

# Builtin number format ids that denote dates or times (same set xlrd uses):
# 14-22 date/time, 27-36 East Asian date, 45-47 elapsed time, 50-58 East
# Asian date variants.
const BUILTIN_DATE_FORMAT_IDS = Set{Int}([14:22; 27:36; 45:47; 50:58])

# Decide whether a custom number format string denotes a date/time. Quoted
# literals, escaped characters, padding/fill markers and bracket sections
# (colors like [Red], conditions like [<=100]) must be ignored; elapsed-time
# tokens like [h] or [ss] count as time. What remains is a date/time format
# iff it contains any of the date/time format codes y, m, d, h or s.
function is_date_format_string(fmt::AbstractString)
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
    return occursin(r"[ymdhs]"i, String(take!(stripped)))
end

function is_date_format(index::Integer, custom_formats::Dict{UInt16,String})
    # A FORMAT record can redefine any index, including builtin ones, so the
    # formats stored in the file take precedence over the builtin table.
    haskey(custom_formats, index) && return is_date_format_string(custom_formats[index])
    return Int(index) in BUILTIN_DATE_FORMAT_IDS
end

const MILLISECONDS_PER_DAY = 86_400_000

"""
    excel_serial_to_temporal(value, is1904)

Convert an Excel serial date/time value to a `DateTime`, or to a `Time` when
the value has no date component (`0 <= value < 1`). In the 1900 date system
the nonexistent date 1900-02-29 (serial 60) maps to 1900-02-28. Negative
values are not valid dates and are returned unchanged as `Float64`.
"""
function excel_serial_to_temporal(value::Float64, is1904::Bool)
    value < 0 && return value
    days = floor(Int, value)
    ms = round(Int, (value - days) * MILLISECONDS_PER_DAY)
    if ms >= MILLISECONDS_PER_DAY
        days += 1
        ms = 0
    end
    t = Time(0) + Millisecond(ms)
    days == 0 && return t
    if is1904
        d = Date(1904, 1, 1) + Day(days)
    elseif days == 60
        d = Date(1900, 2, 28)
    elseif days < 60
        d = Date(1899, 12, 31) + Day(days)
    else
        d = Date(1899, 12, 30) + Day(days)
    end
    return DateTime(d, t)
end
