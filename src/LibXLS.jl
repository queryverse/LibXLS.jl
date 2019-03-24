
module LibXLS

import Dates

# Load libreadstat from our deps.jl
const depsjl_path = joinpath(@__DIR__, "..", "deps", "deps.jl")
if !isfile(depsjl_path)
    error("LibXLS not installed properly, run Pkg.build(\"LibXLS\"), restart Julia and try again")
end
include(depsjl_path)

include("c.jl")
include("types.jl")
include("date.jl")
include("workbook.jl")
include("worksheet.jl")

if Sys.WORD_SIZE == 64
    @assert sizeof(st_sheet) == 16
    @assert sizeof(st_sheet_data) == 16
    @assert sizeof(st_sst) == 32
    @assert sizeof(st_cell_data) == 40
    @assert sizeof(st_cell) == 16
    @assert sizeof(st_row) == 16
    @assert sizeof(st_row_data) == 32
    @assert sizeof(st_colinfo) == 16
    @assert sizeof(st_colinfo_data) == 10
    @assert sizeof(st_xf) == 16
    @assert sizeof(st_xf_data) == 24
    @assert sizeof(st_font_data) == 24
    @assert sizeof(st_font) == 16
    @assert sizeof(st_format_data) == 16
    @assert sizeof(st_format) == 16
end

end
