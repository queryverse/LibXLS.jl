
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
        wb = LibXLS.openxls(fp_book1)
        LibXLS.close(wb)
    end

    @testset "Invalid XLS file" begin
        @test_throws ErrorException LibXLS.openxls(fp_xlsx)
    end
end

@testset "workbook" begin
    LibXLS.openxls(fp_book1) do wb
        @test !LibXLS.is1904(wb)
        @test LibXLS.sheetcount(wb) == 2
        @test LibXLS.sheetnames(wb) == [ "Plan1", "Plan2" ]
        @test LibXLS.sheetname(wb, 1) == "Plan1"
        @test LibXLS.sheetname(wb, 2) == "Plan2"
        @test LibXLS.sheetindex(wb, "Plan1") == 1
        @test LibXLS.sheetindex(wb, "Plan2") == 2

        # returns false, is that really it?
        # @test LibXLS.isvisible(wb, 1))
        # @test LibXLS.isvisible(wb, "Plan1")
    end

    LibXLS.openxls(fp_book1_1904) do wb
        @test LibXLS.is1904(wb)
        @test LibXLS.sheetcount(wb) == 2
        @test LibXLS.sheetnames(wb) == [ "Plan1", "Plan2" ]
    end
end

@testset "worksheet" begin
    LibXLS.openxls(fp_book1) do wb
        ws = wb["Plan2"]
        LibXLS.close(ws)
    end
end
