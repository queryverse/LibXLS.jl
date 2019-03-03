
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

struct st_cell_data
    id::UInt16
    row::UInt16
    col::UInt16
    xf::UInt16
    str::Cstring
    d::Cdouble
    l::UInt32
    width::UInt16
    colspan::UInt16
    rowspan::UInt16
    isHidden::Cuchar
end

struct st_cell
    count::UInt32
    cell::Ptr{st_cell_data}
end

struct st_row_data
    index::UInt16
    fcell::UInt16
    lcell::UInt16
    height::UInt16
    flags::UInt16
    xf::UInt16
    xfflags::Cuchar
    cells::st_cell
end

struct st_row
    lastcol::UInt16 # numCols - 1
    lastrow::UInt16 # numRows - 1
    row::Ptr{st_row_data}
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

@enum XLSRecord::UInt16 begin
    XLS_RECORD_EOF              = 0x000A
    XLS_RECORD_DEFINEDNAME      = 0x0018
    XLS_RECORD_NOTE             = 0x001C
    XLS_RECORD_1904             = 0x0022
    XLS_RECORD_CONTINUE         = 0x003C
    XLS_RECORD_WINDOW1          = 0x003D
    XLS_RECORD_CODEPAGE         = 0x0042
    XLS_RECORD_OBJ              = 0x005D
    XLS_RECORD_MERGEDCELLS      = 0x00E5
    XLS_RECORD_DEFCOLWIDTH      = 0x0055
    XLS_RECORD_COLINFO          = 0x007D
    XLS_RECORD_BOUNDSHEET       = 0x0085
    XLS_RECORD_PALETTE          = 0x0092
    XLS_RECORD_MULRK            = 0x00BD
    XLS_RECORD_MULBLANK         = 0x00BE
    XLS_RECORD_DBCELL           = 0x00D7
    XLS_RECORD_XF               = 0x00E0
    XLS_RECORD_MSODRAWINGGROUP  = 0x00EB
    XLS_RECORD_MSODRAWING       = 0x00EC
    XLS_RECORD_SST              = 0x00FC
    XLS_RECORD_LABELSST         = 0x00FD
    XLS_RECORD_EXTSST           = 0x00FF
    XLS_RECORD_TXO              = 0x01B6
    XLS_RECORD_HYPERREF         = 0x01B8
    XLS_RECORD_BLANK            = 0x0201
    XLS_RECORD_NUMBER           = 0x0203
    XLS_RECORD_LABEL            = 0x0204
    XLS_RECORD_BOOLERR          = 0x0205
    XLS_RECORD_STRING           = 0x0207 # only follows a formula
    XLS_RECORD_ROW              = 0x0208
    XLS_RECORD_INDEX            = 0x020B
    XLS_RECORD_ARRAY            = 0x0221 # Array-entered formula
    XLS_RECORD_DEFAULTROWHEIGHT = 0x0225
    XLS_RECORD_FONT             = 0x0031 # spec says 0x0231 but Excel expects 0x0031
    XLS_RECORD_FONT_ALT         = 0x0231
    XLS_RECORD_WINDOW2          = 0x023E
    XLS_RECORD_RK               = 0x027E
    XLS_RECORD_STYLE            = 0x0293
    XLS_RECORD_FORMULA          = 0x0006
    XLS_RECORD_FORMULA_ALT      = 0x0406 # Apple Numbers bug
    XLS_RECORD_FORMAT           = 0x041E
    XLS_RECORD_BOF              = 0x0809
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
