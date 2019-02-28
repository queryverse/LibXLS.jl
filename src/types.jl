
abstract type AbstractWorkbook end

mutable struct Worksheet{W<:AbstractWorkbook}
    parent::W
    handle::Ptr{xlsWorkSheet}

    function Worksheet(parent::W, handle::Ptr{xlsWorkSheet}) where {W<:AbstractWorkbook}
        new_ws = new{W}(parent, handle)
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
    sheets_info::Vector{WorksheetInfo}
    sheetname_index::Dict{String, Int}
    sheets::Dict{Int, Worksheet}

    function Workbook(handle::Ptr{xlsWorkBook})
        new_wb = new(handle, false, "", Vector{WorksheetInfo}(), Dict{String, Int}(), Dict{Int, Worksheet}())
        finalizer(close, new_wb)

        # parse c struct xlsWorkBook
        wb = unsafe_load(handle)
        new_wb.is1904 = Bool(wb.is1904)
        new_wb.charset = unsafe_string(wb.charset)

        for i in 1:wb.sheets.count
            sheet_data = unsafe_load(wb.sheets.sheet, i)
            push!(new_wb.sheets_info, WorksheetInfo(unsafe_string(sheet_data.name), Bool(sheet_data.visibility)))

            new_wb.sheetname_index[sheetname(new_wb, i)] = i
        end
        return new_wb
    end
end
