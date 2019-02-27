module LibXLS

# Load libreadstat from our deps.jl
const depsjl_path = joinpath(@__DIR__, "..", "deps", "deps.jl")
if !isfile(depsjl_path)
    error("LibXLS not installed properly, run Pkg.build(\"LibXLS\"), restart Julia and try again")
end
include(depsjl_path)

include("c.jl")
include("types.jl")
include("workbook.jl")

end
