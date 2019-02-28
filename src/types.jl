
struct WorksheetInfo
    name::String
    isvisible::Bool
end

mutable struct Workbook
    handle::Ptr{xlsWorkBook}
    is1904::Bool
    charset::String
    sheets_info::Vector{WorksheetInfo}
    sheetname_index::Dict{String, Int}

    function Workbook(handle::Ptr{xlsWorkBook})
        new_xls = new(handle, false, "", Vector{WorksheetInfo}(), Dict{String, Int}())
        finalizer(closexls, new_xls)

        # parse c struct xlsWorkBook
        wb = unsafe_load(handle)
        new_xls.is1904 = Bool(wb.is1904)
        new_xls.charset = unsafe_string(wb.charset)

        for i in 1:wb.sheets.count
            sheet_data = unsafe_load(wb.sheets.sheet, i)
            push!(new_xls.sheets_info, WorksheetInfo(unsafe_string(sheet_data.name), Bool(sheet_data.visibility)))

            new_xls.sheetname_index[sheetname(new_xls, i)] = i
        end
        return new_xls
    end
end
