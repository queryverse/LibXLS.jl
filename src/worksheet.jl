
function close(ws::Worksheet)
    if ws.handle != C_NULL
        xls_close_WS(ws.handle)
        ws.handle = C_NULL
    end
end

function Worksheet(wb::Workbook, sheet_index::Integer)
    check_valid_sheetindex(wb, sheet_index)
    handle = xls_getWorkSheet(wb.handle, sheet_index - 1)
    if handle == C_NULL
        error("Couldn't open Worksheet $sheet_index.")
    end

    expect( xls_parseWorkSheet(handle) , "Failed parsing sheet $sheet_index" )

    # parse c struct xlsWorkSheet
    xlsws = unsafe_load(handle)

    return Worksheet(wb, sheet_index, handle, xlsws.rows)
end

@inline last_row_index(ws::Worksheet) = ws.rows.lastrow + 1
@inline last_column_index(ws::Worksheet) = ws.rows.lastcol + 1

const MAX_WORKSHEET_ROW_INDEX = typemax(UInt16)

# See implementation of `xlsRow *xls_row(xlsWorkSheet* pWS, WORD cellRow)`
@inline is_valid_worksheet_row(ws::Worksheet, row::Integer) = (
    (row <= MAX_WORKSHEET_ROW_INDEX)
    && (0 < row <= last_row_index(ws))
    && ws.rows.row != C_NULL
)

@inline is_valid_worksheet_column(ws::Worksheet, column::Integer) = 0 < column <= last_column_index(ws)

@inline check_valid_worksheet_column(ws::Worksheet, column::Integer) = @assert is_valid_worksheet_column(ws, column) "Worksheet column out of bounds: $column."
@inline check_valid_worksheet_row(ws::Worksheet, row::Integer) = @assert is_valid_worksheet_row(ws, row) "Worksheet Row out of bounds: $row."
Base.size(ws::Worksheet) = ( last_row_index(ws), last_column_index(ws) )

function WorksheetRow(ws::Worksheet, row::Integer)
    check_valid_worksheet_row(ws, row)

    # adds to buffer if not present
    if row ∉ keys(ws.worksheet_rows)
        row_data = unsafe_load(ws.rows.row, row)
        worksheet_row = WorksheetRow(ws, row_data)
        ws.worksheet_rows[row] = worksheet_row
    end

    return ws.worksheet_rows[row]
end

function celldata(ws::Worksheet, row::Integer, col::Integer) :: st_cell_data
    check_valid_worksheet_column(ws, col)
    wsrow = WorksheetRow(ws, row)

    if col ∉ keys(wsrow.cell_data)
        cell_data_ptr = wsrow.row_data.cells.cell
        cell_data = unsafe_load(cell_data_ptr, col)
        wsrow.cell_data[col] = cell_data
    end

    return wsrow.cell_data[col]
end

function Base.getindex(ws::Worksheet, row::Integer, column::Integer)
    cell = celldata(ws, row, column)
    cell_record = XLSRecord(cell.id)

    if cell_record == XLS_RECORD_NUMBER || cell_record == XLS_RECORD_RK
        return cell.d
    elseif cell_record == XLS_RECORD_BLANK
        return missing
    elseif cell_record == XLS_RECORD_LABELSST
        return unsafe_string(cell.str)
    else
        error("Unsupported Record: $cell_record.")
    end
end

sheetindex(ws::Worksheet) = ws.sheet_index
sheetname(ws::Worksheet) = sheetname(ws.parent, sheetindex(ws))
