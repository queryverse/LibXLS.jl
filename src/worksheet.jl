
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

    return Worksheet(handle)
end
