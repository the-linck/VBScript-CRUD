# [VBScript-CRUD](../README.md)
## ADOConstants

* ADODB Type Constants  
    *Mainly used on Parameter types*
    * *int* **adEmpty** *[0]*  
        No value
    * *int* **adSmallInt** *[2]*  
        A 2-byte signed integer.
    * *int* **adInteger** *[3]*  
        A 4-byte signed integer.
    * *int* **adSingle** *[4]*  
        A single-precision floating-point value.
    * *int* **adDouble** *[5]*  
        A double-precision floating-point value.
    * *int* **adCurrency** *[6]*  
        A currency value
    * *int* **adDate** *[7]*  
        The number of days since December 30, 1899 + the fraction of a day.
    * *int* **adBSTR** *[8]*  
        A null-terminated character string.
    * *int* **adIDispatch** *[9]*  
        A pointer to an IDispatch interface on a COM object.  
        Note: Currently not supported by ADO.
    * *int* **adError** *[10]*  
        A 32-bit error code
    * *int* **adBoolean** *[11]*  
        A boolean value.
    * *int* **adVariant** *[12]*  
        An Automation Variant. Note: Currently not supported by ADO.
    * *int* **adIUnknown** *[13]*  
        A pointer to an IUnknown interface on a COM object.  
        Note: Currently not supported by ADO.
    * *int* **adDecimal** *[14]*  
        An exact numeric value with a fixed precision and scale.
    * *int* **adTinyInt** *[16]*  
        A 1-byte signed integer.
    * *int* **adUnsignedTinyInt** *[17]*  
        A 1-byte unsigned integer.
    * *int* **adUnsignedSmallInt** *[18]*  
        A 2-byte unsigned integer.
    * *int* **adUnsignedInt** *[19]*  
        A 4-byte unsigned integer.
    * *int* **adBigInt** *[20]*  
        An 8-byte signed integer.
    * *int* **adUnsignedBigInt** *[21]*  
        An 8-byte unsigned integer.
    * *int* **adFileTime** *[64]*  
        The number of 100-nanosecond intervals since January 1,1601
    * *int* **adGUID** *[72]*  
        A globally unique identifier (GUID)
    * *int* **adBinary** *[128]*  
        A binary value.
    * *int* **adChar** *[129]*  
        A string value.
    * *int* **adWChar** *[130]*  
        A null-terminated Unicode character string.
    * *int* **adNumeric** *[131]*  
        An exact numeric value with a fixed precision and scale.
    * *int* **adUserDefined** *[132]*  
        A user-defined variable.
    * *int* **adDBDate** *[133]*  
        A date value (yyyymmdd).
    * *int* **adDBTime** *[134]*  
        A time value (hhmmss).
    * *int* **adDBTimeStamp** *[135]*  
        A date/time stamp (yyyymmddhhmmss plus a fraction in billionths).
    * *int* **adChapter** *[136]*  
        A 4-byte chapter value that identifies rows in a child rowSet
    * *int* **adPropVariant** *[138]*  
        An Automation PROPVARIANT.
    * *int* **adVarNumeric** *[139]*  
        A numeric value (Parameter object only).
    * *int* **adVarChar** *[200]*  
        A string value (Parameter object only).
    * *int* **adLongVarChar** *[201]*  
        A long string value.
    * *int* **adVarWChar** *[202]*  
        A null-terminated Unicode character string.
    * *int* **adLongVarWChar** *[203]*  
        A long null-terminated Unicode string value.
    * *int* **adVarBinary** *[204]*  
        A binary value (Parameter object only).
    * *int* **adLongVarBinary** *[205]*  
        A long binary value.
* ADODB LockType Constants  
    *Used by Recordset*
    * *int* **adLockUnspecified** *[-1]*  
        Unspecified type of lock. Clones inherits lock type from the original RecordSet.
    * *int* **adLockReadOnly** *[1]*  
        Read-only records
    * *int* **adLockPessimistic** *[2]*  
        Pessimistic locking, record by record. The provider lock records immediately after editing
    * *int* **adLockOptimistic** *[3]*  
        Optimistic locking, record by record. The provider lock records only when calling update
    * *int* **adLockBatchOptimistic** *[4]*  
        Optimistic batch updates. Required for batch update mode
* ADODB Cursor Type Constants  
    *Used by Recordset*
    * *int* **dOpenUnspecified** *[-1]*  
        Does not specify the type of cursor.
    * *int* **adOpenForwardOnly** *[0]*  
        Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a RecordSet.
    * *int* **adOpenKeySet** *[1]*  
        Uses a keySet cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your RecordSet. Data changes by other users are still visible.
    * *int* **adOpenDynamic** *[2]*  
        Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the RecordSet are allowed, except for bookmarks, if the provider doesn't support them.
    * *int* **adOpenStatic** *[3]*  
        Uses a static cursor. A static copy of a Set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
* ADODB Cursor Location Constants  
    *Used by Recordset*
    * *int* **adUseNone** *[1]*  
        OBSOLETE (appears only for backward compatibility). Does not use cursor services
    * *int* **adUseServer** *[2]*  
        Default. Uses a server-side cursor
    * *int* **adUseClient** *[3]*  
        Uses a client-side cursor supplied by a local cursor library. For backward compatibility, the synonym adUseClientBatch is also supported