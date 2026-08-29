@testitem "C struct layout" begin
    # The structs in c.jl mirror libxls's structs byte for byte; a mismatch
    # means unsafe_load reads garbage. These are the sizes of the structs as
    # compiled on all 64-bit platforms (idea from PR #15). 32-bit ABIs differ
    # between operating systems, so no fixed sizes are asserted there.
    if Sys.WORD_SIZE == 64
        @test sizeof(LibXLS.st_sheet_data) == 16
        @test sizeof(LibXLS.st_sheet) == 16
        @test sizeof(LibXLS.st_font_data) == 24
        @test sizeof(LibXLS.st_font) == 16
        @test sizeof(LibXLS.st_format_data) == 16
        @test sizeof(LibXLS.st_format) == 16
        @test sizeof(LibXLS.st_xf_data) == 24
        @test sizeof(LibXLS.st_xf) == 16
        @test sizeof(LibXLS.str_sst_string) == 8
        @test sizeof(LibXLS.st_sst) == 32
        @test sizeof(LibXLS.st_cell_data) == 40
        @test sizeof(LibXLS.st_cell) == 16
        @test sizeof(LibXLS.st_row_data) == 32
        @test sizeof(LibXLS.st_row) == 16
        @test sizeof(LibXLS.st_colinfo_data) == 10
        @test sizeof(LibXLS.st_colinfo) == 16
        @test sizeof(LibXLS.xlsWorkBook) == 168
        @test sizeof(LibXLS.xlsWorkSheet) == 48
    end

    # The version of the loaded C library must be the one the struct mirrors
    # were written against.
    @test startswith(LibXLS.xls_getVersion(), "1.6.")
end

@testitem "Error and format helpers" begin
    for (code, text) in Dict(0x00 => "#NULL!", 0x07 => "#DIV/0!", 0x17 => "#REF!", 0x2A => "#N/A", 0x1D => "#NAME?", 0x24 => "#NUM!", 0x0F => "#VALUE!")
        @test sprint(show, CellError(code)) == text
    end
    @test sprint(show, CellError(0x99)) == "#ERROR(153)!"

    using LibXLS: format_string_kind, format_kind, FORMAT_NONE, FORMAT_DATE, FORMAT_TIME, FORMAT_DATETIME

    @test format_string_kind("yyyy-mm-dd") == FORMAT_DATE
    @test format_string_kind("dd/mm") == FORMAT_DATE
    @test format_string_kind("mmm") == FORMAT_DATE
    @test format_string_kind("[Red]dd/mm/yyyy") == FORMAT_DATE
    @test format_string_kind("h:mm") == FORMAT_TIME
    @test format_string_kind("h:mm AM/PM") == FORMAT_TIME
    @test format_string_kind("hh:mm:ss.000") == FORMAT_TIME
    @test format_string_kind("[h]:mm:ss") == FORMAT_TIME
    @test format_string_kind("mm:ss") == FORMAT_TIME
    @test format_string_kind("yyyy-mm-dd hh:mm") == FORMAT_DATETIME
    @test format_string_kind("d-mmm h:mm AM/PM") == FORMAT_DATETIME
    @test format_string_kind("General") == FORMAT_NONE
    @test format_string_kind("0.00") == FORMAT_NONE
    @test format_string_kind("#,##0.00") == FORMAT_NONE
    @test format_string_kind("0.00E+00") == FORMAT_NONE
    @test format_string_kind("\"m\"0.00") == FORMAT_NONE
    @test format_string_kind("\"unterminated") == FORMAT_NONE
    @test format_string_kind("[Red]0.00") == FORMAT_NONE
    @test format_string_kind("[unterminated") == FORMAT_NONE
    @test format_string_kind("\\y0.00") == FORMAT_NONE
    @test format_string_kind("_y0.00") == FORMAT_NONE
    @test format_string_kind("*y0.00") == FORMAT_NONE

    # builtin format ids, and FORMAT records taking precedence over them
    @test format_kind(14, Dict{UInt16,String}()) == FORMAT_DATE
    @test format_kind(19, Dict{UInt16,String}()) == FORMAT_TIME
    @test format_kind(22, Dict{UInt16,String}()) == FORMAT_DATETIME
    @test format_kind(46, Dict{UInt16,String}()) == FORMAT_TIME
    @test format_kind(0, Dict{UInt16,String}()) == FORMAT_NONE
    @test format_kind(164, Dict{UInt16,String}(0x00a4 => "yyyy-mm-dd")) == FORMAT_DATE
    @test format_kind(14, Dict{UInt16,String}(0x000e => "0.00")) == FORMAT_NONE

    using Dates
    using LibXLS: excel_serial_to_temporal

    @test excel_serial_to_temporal(42066.0, FORMAT_DATE, false) == Date(2015, 3, 3)
    @test excel_serial_to_temporal(42066.0, FORMAT_DATE, false) isa Date
    @test excel_serial_to_temporal(42066.0, FORMAT_DATETIME, false) == DateTime(2015, 3, 3)
    @test excel_serial_to_temporal(42066.0, FORMAT_DATETIME, false) isa DateTime
    @test excel_serial_to_temporal(42039.4263888889, FORMAT_DATETIME, false) == DateTime(2015, 2, 4, 10, 14)
    @test excel_serial_to_temporal(42039.4263888889, FORMAT_DATE, false) == DateTime(2015, 2, 4, 10, 14)
    @test excel_serial_to_temporal(0.626388888888889, FORMAT_TIME, false) == Time(15, 2, 0)
    @test excel_serial_to_temporal(0.626388888888889, FORMAT_DATETIME, false) == Time(15, 2, 0)
    @test excel_serial_to_temporal(1.5416666666, FORMAT_TIME, false) == DateTime(1900, 1, 1, 13, 0)
    @test excel_serial_to_temporal(1.0, FORMAT_DATE, false) == Date(1900, 1, 1)
    @test excel_serial_to_temporal(59.0, FORMAT_DATE, false) == Date(1900, 2, 28)
    @test excel_serial_to_temporal(60.0, FORMAT_DATE, false) == Date(1900, 2, 28) # Excel's phantom 1900-02-29
    @test excel_serial_to_temporal(61.0, FORMAT_DATE, false) == Date(1900, 3, 1)
    @test excel_serial_to_temporal(1.0, FORMAT_DATE, true) == Date(1904, 1, 2)
    @test excel_serial_to_temporal(-1.0, FORMAT_DATE, false) == -1.0 # not a valid date
    # sub-millisecond rounding carries over into the next day
    @test excel_serial_to_temporal(0.99999999999, FORMAT_TIME, false) == DateTime(1900, 1, 1)

    # C error reporting
    @test LibXLS.xls_getError(LibXLS.LIBXLS_OK) isa String
    @test !isempty(LibXLS.xls_getError(LibXLS.LIBXLS_ERROR_PARSE))
    @test LibXLS.expect(LibXLS.LIBXLS_OK, "fine") === nothing
    @test_throws ErrorException LibXLS.expect(LibXLS.LIBXLS_ERROR_READ, "boom")
    err = try
        LibXLS.expect(LibXLS.LIBXLS_ERROR_READ, "boom")
    catch e
        e
    end
    @test occursin("boom", err.msg)
end

@testitem "Corrupt files" begin
    # A file with a valid OLE2 header that libxls cannot parse: the open
    # itself must fail with the error reported by the C library.
    path = joinpath(mktempdir(), "corrupt.xls")
    write(path, vcat(LibXLS.XLS_FILE_HEADER, rand(UInt8, 100)))
    @test_throws ErrorException openxls(path)
end

@testitem "Reading TestData.xls" begin
    using Dates

    filename = normpath(@__DIR__, "TestData.xls")

    @test_throws ErrorException openxls("FileThatDoesNotExist.xls")
    @test_throws ErrorException openxls(normpath(@__DIR__, "runtests.jl"))

    wb = openxls(filename)
    @test sheetcount(wb) == 4
    @test sheetnames(wb) == ["Sheet1", "Second Sheet", "Sheet2", "Empty Sheet"]
    @test LibXLS.sheetindex(wb, "Second Sheet") == 2
    @test_throws ErrorException LibXLS.sheetindex(wb, "No Such Sheet")
    @test !LibXLS.is1904(wb)
    @test LibXLS.isvisible(wb, 1)
    @test LibXLS.isvisible(wb, "Sheet1")
    @test sprint(show, wb) == "LibXLS.Workbook with 4 sheet(s)"

    ws = getworksheet(wb, "Sheet1")
    @test ws === wb["Sheet1"] === wb[1]
    @test size(ws, 1) >= 7
    @test size(ws, 2) >= 14
    @test size(ws) == (size(ws, 1), size(ws, 2))
    @test sprint(show, ws) == "LibXLS.Worksheet Sheet1 ($(size(ws, 1))x$(size(ws, 2)))"

    # Row 3 holds the headers, data starts at row 4; columns C (3) onwards.
    @test ws[1, 1] === missing
    @test ws[3, 3] == "Some Float64s"
    @test ws[4, 3] == 1.0
    @test ws[5, 3] == 1.5
    @test ws[6, 3] == 2.0
    @test ws[4, 4] == "A"
    @test ws[6, 4] == "CCC"
    @test ws[4, 5] === true
    @test ws[5, 5] isa Bool
    @test ws[6, 7] === missing              # "Mixed with NA" NA cell
    @test ws[4, 11] == Date(2015, 3, 3)     # "Some dates"
    @test ws[4, 11] isa Date
    @test ws[5, 11] == DateTime(2015, 2, 4, 10, 14)
    @test ws[5, 11] isa DateTime
    @test ws[6, 11] == Date(1988, 4, 9)
    @test ws[7, 11] == Time(15, 2, 0)
    @test ws[7, 11] isa Time
    @test ws[5, 12] == DateTime(1950, 8, 9, 18, 40)
    @test ws[7, 12] === missing             # "Dates with NA" NA cell
    @test ws[4, 13] isa CellError           # "Some errors"
    @test ws[5, 13] isa CellError
    @test ws[6, 14] isa CellError           # "Errors with NA"
    @test ws[7, 14] === missing

    @test_throws BoundsError ws[0, 1]
    @test_throws BoundsError ws[1, 0]
    @test_throws BoundsError ws[size(ws, 1) + 1, 1]

    ws2 = getworksheet(wb, 2)
    # Data on the second sheet starts at row 8, column 4.
    @test ws2[9, 4] == 1.0
    @test ws2[12, 5] == "CCC"
    @test ws2[10, 6] === false
    @test ws2[13, 9] == Time(15, 2, 0)

    @test isopen(wb)
    close(wb)
    close(wb) # closing twice is fine
    @test !isopen(wb)
    @test_throws ErrorException getworksheet(wb, 1)
    @test_throws ErrorException ws[4, 3]

    # do-block form closes the workbook
    result = openxls(filename) do wb2
        getworksheet(wb2, "Sheet1")[4, 3]
    end
    @test result == 1.0
end

@testitem "Reading book1.xls" begin
    using Dates

    data_folder = normpath(@__DIR__, "..", "data")
    fp_book1 = joinpath(data_folder, "book1.xls")
    fp_book1_1904 = joinpath(data_folder, "book1_1904.xls")
    fp_xlsx = joinpath(data_folder, "blank.xlsx")

    # Checks whether `ws` matches `test_data`, a vector of columns
    function check_test_data(ws, test_data::Vector)
        for col in eachindex(test_data), row in eachindex(test_data[col])
            test_value = test_data[col][row]
            value = ws[row, col]
            if ismissing(test_value) || (isa(test_value, AbstractString) && isempty(test_value))
                @test ismissing(value) || (isa(value, AbstractString) && isempty(value))
            elseif isa(test_value, Float64)
                @test isapprox(value, test_value)
            else
                @test value == test_value
            end
        end
    end

    @test_throws ErrorException openxls(fp_xlsx)

    openxls(fp_book1) do wb
        @test !LibXLS.is1904(wb)
        @test sheetcount(wb) == 2
        @test sheetnames(wb) == ["Plan1", "Plan2"]
        @test LibXLS.sheetname(wb, 1) == "Plan1"
        @test LibXLS.sheetindex(wb, "Plan2") == 2
        @test LibXLS.isvisible(wb, 1)
        @test LibXLS.isvisible(wb, "Plan1")

        @test_throws ErrorException wb["invalid_sheetname"]
        @test_throws ErrorException wb[0]
        @test_throws ErrorException wb[3]

        @test size(wb[1]) == (6, 6)
        @test size(wb[2]) == (4, 5)

        @test LibXLS.sheetindex(wb[1]) == 1
        @test LibXLS.sheetname(wb[1]) == "Plan1"

        ws = wb["Plan1"]
        @test ws[2, 2] == 1
        @test ws[2, 5] isa Date
        check_test_data(ws, [
            [missing for i in 1:6],
            [missing, 1, 2, 3, missing, 5],
            [missing, 1000.1, 1000.2, 1000.3, missing, 1000.5],
            [missing, "abc", "def", "ghi", missing, "xyz"],
            [missing, Date(2018, 12, 1), Date(2018, 12, 31), Date(2019, 1, 1), missing, Date(2019, 2, 26)],
        ])

        ws2 = wb["Plan2"]
        @test ws2[2, 2] == "A"
        check_test_data(ws2, [
            [missing for i in 1:4],
            [missing, "A", "B", "C"],
            ["A", 1, 0.2, 0.3],
            ["B", 0.2, 1, 0.4],
            ["C", 0.3, 0.4, 1],
        ])
    end

    # The 1904 variant holds the same content saved in the 1904 date system,
    # so reading it must produce the same wall-clock dates.
    openxls(fp_book1_1904) do wb
        @test LibXLS.is1904(wb)
        @test sheetnames(wb) == ["Plan1", "Plan2"]
        ws = wb["Plan1"]
        @test ws[2, 5] == Date(2018, 12, 1)
        @test ws[4, 5] == Date(2019, 1, 1)
    end
end
