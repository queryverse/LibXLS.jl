
struct Format
    index::UInt16
    value::String
end

abstract type AbstractWorkbook end
abstract type AbstractWorksheet{W<:AbstractWorkbook} end

struct WorksheetRow{W<:AbstractWorksheet}
    parent::W
    row_data::st_row_data
    cell_data::Dict{UInt16, st_cell_data}

    function WorksheetRow(parent::W, row_data::st_row_data) where {W<:AbstractWorksheet}
        return new{W}(parent, row_data, Dict{UInt16, st_cell_data}())
    end
end

mutable struct Worksheet{W<:AbstractWorkbook} <: AbstractWorksheet{W}
    parent::W
    sheet_index::Int
    handle::Ptr{xlsWorkSheet}
    rows::st_row
    worksheet_rows::Dict{UInt16, WorksheetRow}

    function Worksheet(parent::W, sheet_index::Int, handle::Ptr{xlsWorkSheet}, rows::st_row) where {W<:AbstractWorkbook}
        new_ws = new{W}(parent, sheet_index, handle, rows, Dict{UInt16, WorksheetRow}())
        finalizer(close, new_ws)
        return new_ws
    end
end

struct WorksheetInfo
    name::String
    isvisible::Bool
end

mutable struct Workbook <: AbstractWorkbook
    handle::Ptr{xlsWorkBook}
    is1904::Bool
    charset::String
    xfs::Vector{st_xf_data}
    formats::Vector{Format}
    sheets_info::Vector{WorksheetInfo}
    sheetname_index::Dict{String, Int}
    sheets::Dict{Int, Worksheet}

    function Workbook(handle::Ptr{xlsWorkBook}, is1904::Bool, charset::String, xfs::Vector{st_xf_data}, formats::Vector{Format}, sheets_info::Vector{WorksheetInfo}, sheetname_index::Dict{String, Int}, sheets::Dict{Int, Worksheet})
        new_wb = new(handle, is1904, charset, xfs, formats, sheets_info, sheetname_index, sheets)
        finalizer(close, new_wb)
        return new_wb
    end
end
