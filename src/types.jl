mutable struct XLSWorkBook
    handle::Ptr{xlsWorkBook}

    function XLSWorkBook(handle::Ptr{xlsWorkBook})
        new_wb = new(handle)
        finalizer(closexls, new_wb)
        return new_wb
    end
end
