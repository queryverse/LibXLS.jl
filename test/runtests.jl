
import LibXLS
using Test

const DATA_FOLDER = joinpath(@__DIR__, "..", "data")
@assert isdir(DATA_FOLDER)

fp_book1 = joinpath(DATA_FOLDER, "book1.xls")
@assert isfile(fp_book1)

fp_book1_1904 = joinpath(DATA_FOLDER, "book1_1904.xls")
@assert isfile(fp_book1_1904)

@testset "open/close" begin
    xls = LibXLS.openxls(fp_book1)
    LibXLS.closexls(xls)
end

@testset "is1904" begin
    xls = LibXLS.openxls(fp_book1)
    xls_1904 = LibXLS.openxls(fp_book1_1904)
    @test !LibXLS.is1904(xls)
    @test LibXLS.is1904(xls_1904)
    LibXLS.closexls(xls)
    LibXLS.closexls(xls_1904)
end
