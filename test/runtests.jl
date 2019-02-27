
import LibXLS
using Test

const DATA_FOLDER = joinpath(@__DIR__, "..", "data")
@assert isdir(DATA_FOLDER)

fp_book1 = joinpath(DATA_FOLDER, "book1.xls")
@assert isfile(fp_book1)

@testset "open/close" begin
    xls = LibXLS.openxls(fp_book1)
    LibXLS.closexls(xls)
end
