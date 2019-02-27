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
    codepage::UInt16	# Charset codepage
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
    lastrow::UInt16	# numRows - 1
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
        #     WORD	first;
        #     WORD	last;
        #     WORD	width;
        #     WORD	xf;
        #     WORD	flags;
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

# xlsWorkBook *xls_open_file(const char *file, const char *charset, xls_error_t *outError);
function xls_open_file(file::AbstractString, charset::AbstractString, error_ref::Ref{XLSError})
    ccall((:xls_open_file, libxlsreader), Ptr{Cvoid}, (Cstring, Cstring, Ref{XLSError}), file, charset, error_ref)
end

# void xls_close_WB(xlsWorkBook* pWB);
function xls_close_WB(workbook_handle::Ptr{Cvoid})
    ccall((:xls_close_WB, libxlsreader), Cvoid, (Ptr{Cvoid},), workbook_handle)
end