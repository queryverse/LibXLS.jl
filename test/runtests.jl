
import LibXLS
using Test

const DATA_FOLDER = joinpath(@__DIR__, "..", "data")
@assert isdir(DATA_FOLDER)

fp_book1 = joinpath(DATA_FOLDER, "book1.xls")
@assert isfile(fp_book1)

fp_book1_1904 = joinpath(DATA_FOLDER, "book1_1904.xls")
@assert isfile(fp_book1_1904)

fp_xlsx = joinpath(DATA_FOLDER, "blank.xlsx")
@assert isfile(fp_xlsx)

@testset "open/close" begin
    @testset "Valid XLS file" begin
        xls = LibXLS.openxls(fp_book1)
        LibXLS.closexls(xls)
    end

    @testset "Invalid XLS file" begin
        @test_throws ErrorException LibXLS.openxls(fp_xlsx)
    end
end

@testset "workbook data" begin
    LibXLS.openxls(fp_book1) do xls
        @test !LibXLS.is1904(xls)
        @test LibXLS.sheetcount(xls) == 2
        @test LibXLS.sheetnames(xls) == [ "Plan1", "Plan2" ]
        @test LibXLS.sheetname(xls, 1) == "Plan1"
        @test LibXLS.sheetname(xls, 2) == "Plan2"
        @test LibXLS.sheetindex(xls, "Plan1") == 1
        @test LibXLS.sheetindex(xls, "Plan2") == 2

        # returns false, is that really it?
        # @test LibXLS.isvisible(xls, 1))
        # @test LibXLS.isvisible(xls, "Plan1")
    end

    LibXLS.openxls(fp_book1_1904) do xls
        @test LibXLS.is1904(xls)
        @test LibXLS.sheetcount(xls) == 2
        @test LibXLS.sheetnames(xls) == [ "Plan1", "Plan2" ]
    end
end
