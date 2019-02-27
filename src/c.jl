
# xlsWorkBook *xls_open_file(const char *file, const char *charset, xls_error_t *outError);
function xls_open_file(file::AbstractString, charset::AbstractString, error_ref::Ref{XLSError})
    ccall((:xls_open_file, libxlsreader), Ptr{Cvoid}, (Cstring, Cstring, Ref{XLSError}), file, charset, error_ref)
end

# void xls_close_WB(xlsWorkBook* pWB);
function xls_close_WB(workbook_handle::Ptr{Cvoid})
    ccall((:xls_close_WB, libxlsreader), Cvoid, (Ptr{Cvoid},), workbook_handle)
end
