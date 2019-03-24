
"""
Converts Excel number to Time.
`x` must be between 0 and 1.

To represent Time, Excel uses the decimal part
of a floating point number. `1` equals one day.
"""
function excel_value_to_time(x::Float64) :: Dates.Time
    @assert x >= 0 && x <= 1
    return Dates.Time(Dates.Nanosecond(round(Int, x * 86400) * 1E9 ))
end

time_to_excel_value(x::Dates.Time) :: Float64 = Dates.value(x) / ( 86400 * 1E9 )

"""
Converts Excel number to Date.

See also: [`LibXLS.is1904`](@ref) function.
"""
function excel_value_to_date(x::Int, _is_date_1904::Bool) :: Dates.Date
    if _is_date_1904
        return Dates.Date(Dates.rata2datetime(x + 695056))
    else
        return Dates.Date(Dates.rata2datetime(x + 693594))
    end
end

function date_to_excel_value(date::Dates.Date, _is_date_1904::Bool) :: Int
    if _is_date_1904
        return Dates.datetime2rata(date) - 695056
    else
        return Dates.datetime2rata(date) - 693594
    end
end

"""
Converts Excel number to DateTime.

The decimal part represents the Time (see `_time` function).
The integer part represents the Date.

See also: [`XLSX.isdate1904`](@ref).
"""
function excel_value_to_datetime(x::Float64, _is_date_1904::Bool) :: Dates.DateTime
    @assert x >= 0

    local dt::Dates.Date
    local hr::Dates.Time

    dt_part = trunc(Int, x)
    hr_part = x - dt_part

    dt = excel_value_to_date(dt_part, _is_date_1904)
    hr = excel_value_to_time(hr_part)

    return dt + hr
end

function datetime_to_excel_value(dt::Dates.DateTime, _is_date_1904::Bool) :: Float64
    return date_to_excel_value(Dates.Date(dt), _is_date_1904) + time_to_excel_value(Dates.Time(dt))
end

function parse_excel_date_or_datetime(x::Float64, _is_date_1904::Bool) :: Union{Dates.DateTime, Dates.Date}
	if isinteger(x)
		return excel_value_to_date(Int(x), _is_date_1904)
	else
		return excel_value_to_datetime(x, _is_date_1904)
	end
end
