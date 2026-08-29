function Worksheet(wb::Workbook, sheet_index::Integer)
    check_valid_sheetindex(wb, sheet_index)
    handle = xls_getWorkSheet(wb.handle, sheet_index - 1)
    if handle == C_NULL
        error("Couldn't open worksheet $sheet_index.")
    end

    expect(xls_parseWorkSheet(handle), "Failed parsing sheet $sheet_index")

    xlsws = unsafe_load(handle)

    return Worksheet(wb, sheet_index, handle, Int(xlsws.rows.lastrow) + 1, Int(xlsws.rows.lastcol) + 1)
end

function Base.close(ws::Worksheet)
    if ws.handle != C_NULL
        xls_close_WS(ws.handle)
        ws.handle = C_NULL
    end
    return nothing
end

"""
    size(ws::Worksheet[, dim])

The dimensions of the used cell range of the worksheet, as (rows, columns).
"""
Base.size(ws::Worksheet) = (ws.lastrow, ws.lastcol)
Base.size(ws::Worksheet, dim::Integer) = size(ws)[dim]

sheetindex(ws::Worksheet) = ws.sheet_index
sheetname(ws::Worksheet) = sheetname(ws.parent, sheetindex(ws))

function Base.show(io::IO, ws::Worksheet)
    print(io, "LibXLS.Worksheet $(sheetname(ws)) ($(ws.lastrow)x$(ws.lastcol))")
end

"""
    getindex(ws::Worksheet, row, col)

Return the value of the cell at the (1-based) `row` and `col`. Depending on
the cell this is a `Float64`, `String`, `Bool`, `DateTime`, `Time`,
[`CellError`](@ref) or `missing` (for blank cells).
"""
function Base.getindex(ws::Worksheet, row::Integer, col::Integer)
    ws.handle == C_NULL && error("Worksheet is closed.")
    (1 <= row <= ws.lastrow && 1 <= col <= ws.lastcol) || throw(BoundsError(ws, (row, col)))

    cell_ptr = xls_cell(ws.handle, row - 1, col - 1)
    cell_ptr == C_NULL && return missing
    cell = unsafe_load(cell_ptr)

    id = cell.id
    if id == UInt16(XLS_RECORD_BLANK)
        return missing
    elseif id == UInt16(XLS_RECORD_NUMBER) || id == UInt16(XLS_RECORD_RK)
        return number_value(ws.parent, cell)
    elseif id == UInt16(XLS_RECORD_LABELSST) || id == UInt16(XLS_RECORD_LABEL) || id == UInt16(XLS_RECORD_RSTRING)
        return cell.str == C_NULL ? missing : unsafe_string(cell.str)
    elseif id == UInt16(XLS_RECORD_BOOLERR)
        # libxls stores the bool or error code in `d` and marks which one it
        # is by setting `str` to "bool" or "error".
        if cell_marker(cell) == "error"
            return CellError(UInt8(round(Int, cell.d) & 0xff))
        else
            return cell.d != 0
        end
    elseif id == UInt16(XLS_RECORD_FORMULA) || id == UInt16(XLS_RECORD_FORMULA_ALT)
        # For a numeric formula result libxls sets `l` to 0 and `d` to the
        # value; otherwise `l` is 0xffff and `str` marks the result as
        # "bool" or "error" (again with the value in `d`), or holds the
        # string result itself.
        if cell.l == 0
            return number_value(ws.parent, cell)
        else
            marker = cell_marker(cell)
            marker == "bool" && return cell.d != 0
            marker == "error" && return CellError(UInt8(round(Int, cell.d) & 0xff))
            return marker
        end
    else
        error("Unsupported cell record 0x$(string(id, base = 16, pad = 4)) at ($row, $col).")
    end
end

cell_marker(cell::st_cell_data) = cell.str == C_NULL ? "" : unsafe_string(cell.str)

function number_value(wb::Workbook, cell::st_cell_data)
    xf_index = Int(cell.xf) + 1
    if xf_index <= length(wb.xf_kind)
        kind = wb.xf_kind[xf_index]
        if kind != FORMAT_NONE
            return excel_serial_to_temporal(cell.d, kind, wb.is1904)
        end
    end
    return cell.d
end
