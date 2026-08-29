module LibXLS

using Dates
using libxls_jll: libxlsreader

export openxls, sheetcount, sheetnames, getworksheet, CellError

include("c.jl")
include("types.jl")
include("formats.jl")
include("workbook.jl")
include("worksheet.jl")

end
