
function openxls(filename::AbstractString) :: XLSWorkBook
    @assert isfile(filename) "$filename not found."
    error_ref = Ref{XLSError}()
    handle = xls_open_file(filename, "UTF-8", error_ref)
    if handle == C_NULL
        throw(error_ref[])
    end
    return XLSWorkBook(handle)
end

function closexls(xls::XLSWorkBook)
    if xls.handle != C_NULL
        xls_close_WB(xls.handle)
        xls.handle = C_NULL
    end
end
