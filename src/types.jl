
@enum XLSError::UInt32 begin
    LIBXLS_OK           = 0
    LIBXLS_ERROR_OPEN   = 1
    LIBXLS_ERROR_SEEK   = 2
    LIBXLS_ERROR_READ   = 3
    LIBXLS_ERROR_PARSE  = 4
    LIBXLS_ERROR_MALLOC = 5
end

mutable struct XLSWorkBook
    handle::Ptr{Cvoid}

    function XLSWorkBook(handle::Ptr{Cvoid})
        new_wb = new(handle)
        finalizer(closexls, new_wb)
        return new_wb
    end
end
