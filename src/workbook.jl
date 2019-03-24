
function Workbook(filepath::AbstractString)

    # get a workbook handle
    check_xls_file_format(filepath)
    error_ref = Ref{XLSError}()
    handle = xls_open_file(filepath, "UTF-8", error_ref)
    if handle == C_NULL
        err = error_ref[]
        @assert err != LIBXLS_OK # shouldn't happen
        expect(err, "Error opening $filepath")
    end

    # creates workbook struct
    new_wb = Workbook(handle, false, "", Vector{st_xf_data}(), Vector{Format}(), Vector{WorksheetInfo}(), Dict{String, Int}(), Dict{Int, Worksheet}())

    # parse c struct xlsWorkBook to Workbook
    xlswb = unsafe_load(handle)
    new_wb.is1904 = Bool(xlswb.is1904)
    new_wb.charset = unsafe_string(xlswb.charset)
    for i in 1:xlswb.sheets.count
        sheet_data = unsafe_load(xlswb.sheets.sheet, i)
        push!(new_wb.sheets_info, WorksheetInfo(unsafe_string(sheet_data.name), Bool(sheet_data.visibility)))

        new_wb.sheetname_index[sheetname(new_wb, i)] = i
    end

    for i in 1:xlswb.xfs.count
        xf_data = unsafe_load(xlswb.xfs.xf_data, i)
        push!(new_wb.xfs, xf_data)
    end

    return new_wb
end

openxls(filepath::AbstractString) :: Workbook = Workbook(filepath)

function openxls(f::Function, filepath::AbstractString)
    wb = openxls(filepath)

    try
        f(wb)
    finally
        close(wb)
    end
end

const XLS_FILE_HEADER = [ 0xd0, 0xcf, 0x11, 0xe0 ]
const ZIP_FILE_HEADER = [ 0x50, 0x4b, 0x03, 0x04 ]

function check_xls_file_format(filepath::AbstractString)
    @assert isfile(filepath) "File $filepath not found."

    local header::Vector{UInt8}

    open(filepath, "r") do io
        header = Base.read(io, 4)
    end

    if header == XLS_FILE_HEADER
        return
    elseif header == ZIP_FILE_HEADER
        error("$filepath is either an Excel file in the new XLSX format, or a Zip file. This package does not support XLSX file format.")
    else
        error("$filepath is not a valid XLS file.")
    end
end

function close(wb::Workbook)
    if wb.handle != C_NULL
        xls_close_WB(wb.handle)
        wb.handle = C_NULL
    end
end

sheetcount(wb::Workbook) :: Int = length(wb.sheets_info)
is1904(wb::Workbook) :: Bool = wb.is1904
sheetname(wb::Workbook, sheet_index::Integer) :: String = wb.sheets_info[sheet_index].name

@inline is_valid_sheetindex(wb::Workbook, sheet_index::Integer) = 0 < sheet_index <= sheetcount(wb)
@inline check_valid_sheetindex(wb::Workbook, sheet_index::Integer) = @assert is_valid_sheetindex(wb, sheet_index) "$sheet_index is not a valid sheet index."
@inline is_valid_sheetname(wb::Workbook, sheet_name::AbstractString) = sheet_name ∈ keys(wb.sheetname_index)
@inline check_valid_sheetname(wb::Workbook, sheet_name::AbstractString) = @assert is_valid_sheetname(wb, sheet_name) "$sheet_name is not a valid sheet name."

@inline function sheetindex(wb::Workbook, sheet_name::AbstractString) :: Int
    check_valid_sheetname(wb, sheet_name)
    return wb.sheetname_index[sheet_name]
end

sheetnames(wb::Workbook) :: Vector{String} = [ sheetname(wb, i) for i in 1:sheetcount(wb) ]
isvisible(wb::Workbook, sheet_index::Integer) :: Bool = wb.sheets_info[sheet_index].isvisible
isvisible(wb::Workbook, sheet_name::AbstractString) :: Bool = isvisible(wb, sheetindex(wb, sheet_name))

function getworksheet(wb::Workbook, sheet_index::Integer) :: Worksheet
    if sheet_index ∉ keys(wb.sheets)
        # add new sheet to buffer
        wb.sheets[sheet_index] = Worksheet(wb, sheet_index)
    end
    return wb.sheets[sheet_index]
end

getworksheet(wb::Workbook, sheet_name::AbstractString) :: Worksheet = getworksheet(wb, sheetindex(wb, sheet_name))

Base.getindex(wb::Workbook, sheet_index::Integer) = getworksheet(wb, sheet_index)
Base.getindex(wb::Workbook, sheet_name::AbstractString) = getworksheet(wb, sheet_name)
