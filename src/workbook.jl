
function openxls(filepath::AbstractString) :: Workbook
    check_xls_file_format(filepath)
    error_ref = Ref{XLSError}()
    handle = xls_open_file(filepath, "UTF-8", error_ref)
    if handle == C_NULL
        throw(error_ref[])
    end
    return Workbook(handle)
end

function openxls(f::Function, filepath::AbstractString)
    xls = openxls(filepath)

    try
        f(xls)
    finally
        closexls(xls)
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

function closexls(xls::Workbook)
    if xls.handle != C_NULL
        xls_close_WB(xls.handle)
        xls.handle = C_NULL
    end
end

sheetcount(xls::Workbook) :: Int = length(xls.sheets_info)
is1904(xls::Workbook) :: Bool = xls.is1904
sheetname(xls::Workbook, sheet_index::Integer) :: String = xls.sheets_info[sheet_index].name

function sheetindex(xls::Workbook, sheet_name::AbstractString) :: Int
    @assert sheet_name ∈ keys(xls.sheetname_index) "$sheet_name is not a valid sheet name."
    return xls.sheetname_index[sheet_name]
end

sheetnames(xls::Workbook) :: Vector{String} = [ sheetname(xls, i) for i in 1:sheetcount(xls) ]
isvisible(xls::Workbook, sheet_index::Integer) :: Bool = xls.sheets_info[sheet_index].isvisible
isvisible(xls::Workbook, sheet_name::AbstractString) :: Bool = isvisible(xls, sheetindex(xls, sheet_name))
