
function openxls(filename::AbstractString) :: XLSWorkbook
    @assert isfile(filename) "$filename not found."
    error_ref = Ref{XLSError}()
    handle = xls_open_file(filename, "UTF-8", error_ref)
    if handle == C_NULL
        throw(error_ref[])
    end
    return XLSWorkbook(handle)
end

function openxls(f::Function, filename::AbstractString)
    xls = openxls(filename)

    try
        f(xls)
    finally
        closexls(xls)
    end
end

function closexls(xls::XLSWorkbook)
    if xls.handle != C_NULL
        xls_close_WB(xls.handle)
        xls.handle = C_NULL
    end
end

sheetcount(xls::XLSWorkbook) = length(xls.sheets_info)
is1904(xls::XLSWorkbook) = xls.is1904
sheetname(xls::XLSWorkbook, sheet_index::Integer) = xls.sheets_info[sheet_index].name

function sheetindex(xls::XLSWorkbook, sheet_name::AbstractString) :: Int
    @assert sheet_name ∈ keys(xls.sheetname_index) "$sheet_name is not a valid sheet name."
    return xls.sheetname_index[sheet_name]
end

sheetnames(xls::XLSWorkbook) = [ sheetname(xls, i) for i in 1:sheetcount(xls) ]
isvisible(xls::XLSWorkbook, sheet_index::Integer) = xls.sheets_info[sheet_index].isvisible
isvisible(xls::XLSWorkbook, sheet_name::AbstractString) = isvisible(xls, sheetindex(xls, sheet_name))
