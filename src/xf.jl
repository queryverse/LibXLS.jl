
xf(ws::Worksheet, cell::st_cell_data) = ws.parent.xfs[cell.xf + 1]
xf(ws::Worksheet, row::Integer, col::Integer) = xf(ws, celldata(ws, row, col))

#=
number_format(ws::Worksheet, cell::st_cell_data) = number_format(ws, ws.parent.xfs[cell.xf + 1])

function number_format(ws::Worksheet, cell_xf::st_xf_data)
    println("inspecting cell_xf $cell_xf.")

    if is_cell_xf(cell_xf)
        # cell xf
        #     In cell XFs a cleared bit means the attributes of the
        #     parent style XF are used (but only if the attributes are valid there),
        #     a set bit means the attributes of this XF are used.

        if !hasmask( cell_xf.usedattr, 0x01 )
            # number format bit is cleared
            # will use parent's style
            return number_format(ws, parent_xf(ws, cell_xf))
        else
            # number format bit is set. Will use style from the current xf.
            return cellformat(ws, cell_xf)
        end
    else
        # style xf
        if !hasmask(cell_xf.usedattr, 0x01)
            # format is valid
            return cellformat(ws, cell_xf)
        else
            # format is not valid and should be ignored
            return nothing
        end
    end
end

is_style_xf(xf::st_xf_data) = hasmask(xf.type, 0x0004)
is_cell_xf(xf::st_xf_data) = !is_style_xf(xf)

function parent_xf(ws::Worksheet, xf::st_xf_data)
    parent_xf_index = xf.type >> 4
    @assert parent_xf_index != 4095 "current xf does not have a parent xf."
    return ws.parent.xfs[parent_xf_index + 1] # xf_index is 0-based
end

function cellformat(ws::Worksheet, xf::st_xf_data)
    format_index = xf.format

    for format in ws.parent.formats
        if format.index == format_index
            return format
        end
    end

    return format_index
end

function cellformat(ws::Worksheet, cell::st_cell_data)
    wb = ws.parent
    xf_index = cell.xf
    xf_data = wb.xfs[xf_index + 1] # xf_index is 0-based
    return cellformat(ws, xf_data)
end

cellformat(ws::Worksheet, row::Integer, column::Integer) = cellformat(ws, celldata(ws, row, column))
=#

function is_date_format(ws::Worksheet, cell::st_cell_data) :: Bool
    cell_xf = xf(ws, cell)
    format = cell_xf.format
    if (14 <= format <= 22) || ( 45 <= format <= 47)
        return true
    end

    return false
end
