
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

# Checks wether `matrix` equals `test_data`, where test_data is a vector of columns
function check_test_data(matrix, test_data::Vector)

    function size_of_data(d::Vector)
        isempty(d) && return (0, 0)
        return length(d[1]), length(d)
    end

    rows, cols = size_of_data(test_data)

    for row in 1:rows, col in 1:cols
        test_value = test_data[col][row]
        value = matrix[row, col]
        if ismissing(test_value) || ( isa(test_value, AbstractString) && isempty(test_value) )
            @test ismissing(value) || ( isa(value, AbstractString) && isempty(value) )
        else
            if isa(test_value, Float64)
                @test isapprox(value, test_value)
            else
                @test value == test_value
            end
        end
    end

    nothing
end

function debug_cell_data(cell::LibXLS.st_cell_data)
    println("Cell Data Debug")
    println("record = ", cell.id)
    println("xf = ", cell.xf)
    if cell.str != C_NULL
        println("str = ", unsafe_string(cell.str))
    end

    println("d = ", cell.d)
    println("l = ", Int(cell.l))
end

@testset "Workbook" begin

    @testset "Valid XLS file" begin
        wb = LibXLS.openxls(fp_book1)
        LibXLS.close(wb)
    end

    @testset "Invalid XLS file" begin
        @test_throws ErrorException LibXLS.openxls(fp_xlsx)
    end

    @testset "Workbook info" begin
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
end

@testset "Worksheet" begin
    LibXLS.openxls(fp_book1) do wb
        @testset "open/close" begin
            let
                ws = wb["Plan2"]

                # closing ws will cause segfault or errors
                # on the remaining tests
                # LibXLS.close(ws)
            end

            let
                ws = wb["Plan1"]
                ws = wb[1]
                ws = wb[2]
            end
        end

        @testset "bounds" begin
            @test_throws AssertionError wb["invalid_sheetname"]
            @test_throws AssertionError wb[0]
            @test_throws AssertionError wb[3]
        end

        @testset "size" begin
            @test size(wb[1]) == (6, 7)
            @test size(wb[2]) == (4, 6)
        end

        @testset "sheetname" begin
            ws = wb[1]
            @test LibXLS.sheetindex(ws) == 1
            @test LibXLS.sheetname(ws) == "Plan1"
        end

        @testset "row data" begin
            let
                ws = wb["Plan1"]
                @test ws[2, 2] == 1

                test_data = [ [missing for i in 1:6],
                              [missing, 1, 2, 3, missing, 5],
                              [missing, 1000.1, 1000.2, 1000.3, missing, 1000.5],
                              [missing, "abc", "def", "ghi", missing, "xyz"],
                              # [missing, Date(2018, 12, 1), Date(2018, 12, 31), Date(2019, 1, 1), missing, Date(2019, 2, 26)]
                            ]
                check_test_data(ws, test_data)
            end

            let
                ws = wb["Plan2"]
                @test ws[2,2] == "A"

                test_data = [ [missing for i in 1:4],
                              [missing, "A", "B", "C"],
                              ["A", 1, 0.2, 0.3],
                              ["B", 0.2, 1, 0.4],
                              ["C", 0.3, 0.4, 1]
                            ]
                check_test_data(ws, test_data)
            end
        end
    end

    LibXLS.openxls(fp_book1) do wb
        ws = wb["Plan2"]
        @test ws[2,2] == "A"
    end

    LibXLS.openxls(fp_book1) do wb
        ws1 = wb["Plan1"]
        ws2 = wb["Plan2"]
        @test ws1[2,2] == 1
        @test ws2[2,2] == "A"
    end
end
