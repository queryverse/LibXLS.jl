abstract type AbstractWorkbook end

struct WorksheetInfo
    name::String
    isvisible::Bool
end

"""
    CellError

An Excel cell that holds an Excel error value such as `#DIV/0!` or `#N/A`.
`code` is the BIFF error code.
"""
struct CellError
    code::UInt8
end

const EXCEL_ERROR_STRINGS = Dict{UInt8,String}(
    0x00 => "#NULL!",
    0x07 => "#DIV/0!",
    0x0F => "#VALUE!",
    0x17 => "#REF!",
    0x1D => "#NAME?",
    0x24 => "#NUM!",
    0x2A => "#N/A",
)

function Base.show(io::IO, e::CellError)
    print(io, get(EXCEL_ERROR_STRINGS, e.code, "#ERROR($(Int(e.code)))!"))
end

mutable struct Worksheet{W <: AbstractWorkbook}
    parent::W
    sheet_index::Int
    handle::Ptr{xlsWorkSheet}
    lastrow::Int # 1-based index of the last row
    lastcol::Int # 1-based index of the last column

    function Worksheet(parent::W, sheet_index::Integer, handle::Ptr{xlsWorkSheet}, lastrow::Integer, lastcol::Integer) where {W <: AbstractWorkbook}
        new_ws = new{W}(parent, Int(sheet_index), handle, Int(lastrow), Int(lastcol))
        finalizer(close, new_ws)
        return new_ws
    end
end

mutable struct Workbook <: AbstractWorkbook
    handle::Ptr{xlsWorkBook}
    is1904::Bool
    charset::String
    sheets_info::Vector{WorksheetInfo}
    sheetname_index::Dict{String,Int}
    sheets::Dict{Int,Worksheet}
    xf_kind::Vector{CellFormatKind} # per XF record: does its number format denote a date and/or time?

    function Workbook(handle::Ptr{xlsWorkBook}, is1904::Bool, charset::String, sheets_info::Vector{WorksheetInfo}, sheetname_index::Dict{String,Int}, sheets::Dict{Int,Worksheet}, xf_kind::Vector{CellFormatKind})
        new_wb = new(handle, is1904, charset, sheets_info, sheetname_index, sheets, xf_kind)
        finalizer(close, new_wb)
        return new_wb
    end
end
