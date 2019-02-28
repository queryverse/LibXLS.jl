
struct st_sheet_data
    filepos::UInt32
    visibility::Cuchar
    type::Cuchar
    name::Cstring
end

struct st_sheet
    count::UInt32 # Count of sheets
    sheet::Ptr{st_sheet_data}
end

struct xlsWorkBook
    olestr::Ptr{Nothing}
    filepos::Int32 # position in file

    # From Header (BIFF)
    is5ver::Cuchar
    is1904::Cuchar
    type::UInt16
    activeSheetIdx:: UInt16 # index of the active sheet

    # Other data
    codepage::UInt16    # Charset codepage
    charset::Cstring
    sheets::st_sheet
    # sst::st_sst # SST table
    # xfs::st_xf # XF table
    # fonts::st_font
    # formats::st_format # FORMAT table

    # summary::Cstring # ole file
    # docSummary::Cstring # ole file
end

struct xls_summaryInfo
    title::Cstring
    subject::Cstring
    author::Cstring
    keywords::Cstring
    comment::Cstring
    lastAuthor::Cstring
    appName::Cstring
    category::Cstring
    manager::Cstring
    company::Cstring
end

struct st_row
    lastcol::UInt16 # numCols - 1
    lastrow::UInt16 # numRows - 1
    row::Ptr{Nothing}
        # struct st_row_data
        # {
        #     WORD index;
        #     WORD fcell;
        #     WORD lcell;
        #     WORD height;
        #     WORD flags;
        #     WORD xf;
        #     BYTE xfflags;
        #     st_cell cells;
        # }
        # * row;
end

struct st_colinfo
    count::UInt32 # Count of COLINFO
    col::Ptr{Nothing}
        # struct st_colinfo_data
        # {
        #     WORD  first;
        #     WORD  last;
        #     WORD  width;
        #     WORD  xf;
        #     WORD  flags;
        # }
        # * col;
end

struct xlsWorkSheet
    filepos::UInt32
    defcolwidth::UInt16
    rows::st_row
    workbook::Ptr{xlsWorkBook}
    colinfo::st_colinfo
end

@enum XLSError::UInt32 begin
    LIBXLS_OK           = 0
    LIBXLS_ERROR_OPEN   = 1
    LIBXLS_ERROR_SEEK   = 2
    LIBXLS_ERROR_READ   = 3
    LIBXLS_ERROR_PARSE  = 4
    LIBXLS_ERROR_MALLOC = 5
end

@inline function expect(err::XLSError, msg::AbstractString)

    local err_str::String = "unknown"

    if err == LIBXLS_OK
        return
    elseif err == LIBXLS_ERROR_OPEN
        err_str = "OPEN"
    elseif err == LIBXLS_ERROR_SEEK
        err_str = "SEEK"
    elseif err == LIBXLS_ERROR_READ
        err_str = "READ"
    elseif err == LIBXLS_ERROR_PARSE
        err_str = "PARSE"
    elseif err == LIBXLS_ERROR_MALLOC
        err_str = "MALLOC"
    end

    error(msg * " (operation $err_str)")
end

# xlsWorkBook *xls_open_file(const char *file, const char *charset, xls_error_t *outError);
function xls_open_file(file::AbstractString, charset::AbstractString, error_ref::Ref{XLSError})
    ccall((:xls_open_file, libxlsreader), Ptr{xlsWorkBook}, (Cstring, Cstring, Ref{XLSError}), file, charset, error_ref)
end

# void xls_close_WB(xlsWorkBook* pWB);
function xls_close_WB(workbook_handle::Ptr{xlsWorkBook})
    ccall((:xls_close_WB, libxlsreader), Cvoid, (Ptr{xlsWorkBook},), workbook_handle)
end

# xlsWorkSheet * xls_getWorkSheet(xlsWorkBook* pWB,int num);
function xls_getWorkSheet(workbook_handle::Ptr{xlsWorkBook}, num::Integer)
    ret = ccall((:xls_getWorkSheet, libxlsreader), Ptr{xlsWorkSheet}, (Ptr{xlsWorkBook}, Cint), workbook_handle, num)
end

# void xls_close_WS(xlsWorkSheet* pWS);
function xls_close_WS(worksheet_handle::Ptr{xlsWorkSheet})
    ccall((:xls_close_WS, libxlsreader), Cvoid, (Ptr{xlsWorkSheet},), worksheet_handle)
end

# xls_error_t xls_parseWorkSheet(xlsWorkSheet* pWS);
function xls_parseWorkSheet(worksheet_handle::Ptr{xlsWorkSheet})
    ccall((:xls_parseWorkSheet, libxlsreader), XLSError, (Ptr{xlsWorkSheet},), worksheet_handle)
end

# xlsSummaryInfo *xls_summaryInfo(xlsWorkBook* pWB);
function xls_summaryInfo(wb)
    ret = ccall((:xls_summaryInfo, libxlsreader), Ptr{xls_summaryInfo}, (Ptr{xlsWorkBook},), wb)
end

# void xls_close_summaryInfo(xlsSummaryInfo *pSI);
function xls_close_summaryInfo(si)
    ccall((:xls_close_summaryInfo, libxlsreader), Cvoid, (Ptr{xls_summaryInfo},), si)
end
