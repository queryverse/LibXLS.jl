function Workbook(filepath::AbstractString)
    check_xls_file_format(filepath)

    error_ref = Ref{XLSError}(LIBXLS_OK)
    handle = xls_open_file(filepath, "UTF-8", error_ref)
    if handle == C_NULL
        expect(error_ref[], "Error opening $filepath")
        error("Error opening $filepath.")
    end

    xlswb = unsafe_load(handle)

    sheets_info = Vector{WorksheetInfo}()
    sheetname_index = Dict{String,Int}()
    for i in 1:xlswb.sheets.count
        sheet_data = unsafe_load(xlswb.sheets.sheet, i)
        name = sheet_data.name == C_NULL ? "" : unsafe_string(sheet_data.name)
        # In BOUNDSHEET records visibility 0 means visible, 1 hidden, 2 very hidden.
        push!(sheets_info, WorksheetInfo(name, sheet_data.visibility == 0))
        sheetname_index[name] = i
    end

    custom_formats = Dict{UInt16,String}()
    for i in 1:xlswb.formats.count
        format_data = unsafe_load(xlswb.formats.format, i)
        if format_data.value != C_NULL
            custom_formats[format_data.index] = unsafe_string(format_data.value)
        end
    end

    xf_kind = Vector{CellFormatKind}(undef, xlswb.xfs.count)
    for i in 1:xlswb.xfs.count
        xf_data = unsafe_load(xlswb.xfs.xf, i)
        xf_kind[i] = format_kind(xf_data.format, custom_formats)
    end

    charset = xlswb.charset == C_NULL ? "" : unsafe_string(xlswb.charset)

    return Workbook(handle, xlswb.is1904 != 0, charset, sheets_info, sheetname_index, Dict{Int,Worksheet}(), xf_kind)
end

"""
    openxls(filepath) -> Workbook
    openxls(f::Function, filepath)

Open the legacy xls file at `filepath` and return a `Workbook`. The second
form calls `f` on the workbook and closes it afterwards.
"""
openxls(filepath::AbstractString)::Workbook = Workbook(filepath)

function openxls(f::Function, filepath::AbstractString)
    wb = openxls(filepath)

    try
        return f(wb)
    finally
        close(wb)
    end
end

const XLS_FILE_HEADER = [0xd0, 0xcf, 0x11, 0xe0]
const ZIP_FILE_HEADER = [0x50, 0x4b, 0x03, 0x04]

function check_xls_file_format(filepath::AbstractString)
    isfile(filepath) || error("File $filepath not found.")

    local header::Vector{UInt8}

    open(filepath, "r") do io
        header = Base.read(io, 4)
    end

    if header == XLS_FILE_HEADER
        return
    elseif header == ZIP_FILE_HEADER
        error("$filepath is either an Excel file in the new XLSX format, or a Zip file. This package does not support the XLSX file format.")
    else
        error("$filepath is not a valid XLS file.")
    end
end

"""
    close(wb::Workbook)

Close the workbook and release the resources held by the C library. Closing
an already closed workbook does nothing. Worksheets obtained from the
workbook must not be accessed afterwards.
"""
function Base.close(wb::Workbook)
    if wb.handle != C_NULL
        for ws in values(wb.sheets)
            close(ws)
        end
        xls_close_WB(wb.handle)
        wb.handle = C_NULL
    end
    return nothing
end

"""
    isopen(wb::Workbook)

Whether the workbook has not been closed yet.
"""
Base.isopen(wb::Workbook) = wb.handle != C_NULL

"""
    sheetcount(wb::Workbook)

The number of sheets in the workbook, including hidden and empty ones.
"""
sheetcount(wb::Workbook)::Int = length(wb.sheets_info)
"""
    is1904(wb::Workbook)

Whether the workbook uses the 1904 date system (the default on classic Mac
versions of Excel) rather than the 1900 date system. Cell values already
account for this, so this is informational only.
"""
is1904(wb::Workbook)::Bool = wb.is1904
"""
    sheetname(wb::Workbook, sheet_index)
    sheetname(ws::Worksheet)

The name of the sheet with the given (1-based) index, or of the given
worksheet.
"""
sheetname(wb::Workbook, sheet_index::Integer)::String = wb.sheets_info[sheet_index].name

@inline is_valid_sheetindex(wb::Workbook, sheet_index::Integer) = 0 < sheet_index <= sheetcount(wb)
@inline function check_valid_sheetindex(wb::Workbook, sheet_index::Integer)
    is_valid_sheetindex(wb, sheet_index) || error("$sheet_index is not a valid sheet index.")
end
@inline is_valid_sheetname(wb::Workbook, sheet_name::AbstractString) = haskey(wb.sheetname_index, sheet_name)
@inline function check_valid_sheetname(wb::Workbook, sheet_name::AbstractString)
    is_valid_sheetname(wb, sheet_name) || error("$sheet_name is not a valid sheet name.")
end

"""
    sheetindex(wb::Workbook, sheet_name)
    sheetindex(ws::Worksheet)

The (1-based) index of the sheet with the given name, or of the given
worksheet.
"""
@inline function sheetindex(wb::Workbook, sheet_name::AbstractString)::Int
    check_valid_sheetname(wb, sheet_name)
    return wb.sheetname_index[sheet_name]
end

"""
    sheetnames(wb::Workbook)

The names of all sheets in the workbook, in sheet order.
"""
sheetnames(wb::Workbook)::Vector{String} = [sheetname(wb, i) for i in 1:sheetcount(wb)]
"""
    isvisible(wb::Workbook, sheet_index_or_name)

Whether the sheet with the given index or name is visible (not hidden).
"""
isvisible(wb::Workbook, sheet_index::Integer)::Bool = wb.sheets_info[sheet_index].isvisible
isvisible(wb::Workbook, sheet_name::AbstractString)::Bool = isvisible(wb, sheetindex(wb, sheet_name))

"""
    getworksheet(wb::Workbook, sheet_index_or_name) -> Worksheet

Return the worksheet with the given (1-based) index or name, parsing it on
first access. `wb[i]` and `wb["name"]` are shorthands for this function.
"""
function getworksheet(wb::Workbook, sheet_index::Integer)::Worksheet
    wb.handle == C_NULL && error("Workbook is closed.")
    if sheet_index ∉ keys(wb.sheets)
        wb.sheets[sheet_index] = Worksheet(wb, sheet_index)
    end
    return wb.sheets[sheet_index]
end

getworksheet(wb::Workbook, sheet_name::AbstractString)::Worksheet = getworksheet(wb, sheetindex(wb, sheet_name))

Base.getindex(wb::Workbook, sheet_index::Integer) = getworksheet(wb, sheet_index)
Base.getindex(wb::Workbook, sheet_name::AbstractString) = getworksheet(wb, sheet_name)

function Base.show(io::IO, wb::Workbook)
    print(io, "LibXLS.Workbook with $(sheetcount(wb)) sheet(s)")
end
