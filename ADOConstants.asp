<%
' ADODB Type Constants
    adEmpty             = 0' No value
    adSmallInt          = 2' A 2-byte signed integer.
    adInteger           = 3' A 4-byte signed integer.
    adSingle            = 4' A single-precision floating-point value.
    adDouble            = 5' A double-precision floating-point value.
    adCurrency          = 6' A currency value
    adDate              = 7' The number of days since December 30, 1899 + the fraction of a day.
    adBSTR              = 8' A null-terminated character string.
    adIDispatch         = 9' A pointer to an IDispatch interface on a COM object. Note: Currently not supported by ADO.
    adError             = 10' A 32-bit error code
    adBoolean           = 11' A boolean value.
    adVariant           = 12' An Automation Variant. Note: Currently not supported by ADO.
    adIUnknown          = 13' A pointer to an IUnknown interface on a COM object. Note: Currently not supported by ADO.
    adDecimal           = 14' An exact numeric value with a fixed precision and scale.
    adTinyInt           = 16' A 1-byte signed integer.
    adUnsignedTinyInt   = 17' A 1-byte unsigned integer.
    adUnsignedSmallInt  = 18' A 2-byte unsigned integer.
    adUnsignedInt       = 19' A 4-byte unsigned integer.
    adBigInt            = 20' An 8-byte signed integer.
    adUnsignedBigInt    = 21' An 8-byte unsigned integer.
    adFileTime          = 64' The number of 100-nanosecond intervals since January 1,1601
    adGUID              = 72' A globally unique identifier (GUID)
    adBinary            = 128' A binary value.
    adChar              = 129' A string value.
    adWChar             = 130' A null-terminated Unicode character string.
    adNumeric           = 131' An exact numeric value with a fixed precision and scale.
    adUserDefined       = 132' A user-defined variable.
    adDBDate            = 133' A date value (yyyymmdd).
    adDBTime            = 134' A time value (hhmmss).
    adDBTimeStamp       = 135' A date/time stamp (yyyymmddhhmmss plus a fraction in billionths).
    adChapter           = 136' A 4-byte chapter value that identifies rows in a child rowSet
    adPropVariant       = 138' An Automation PROPVARIANT.
    adVarNumeric        = 139' A numeric value (Parameter object only).
    adVarChar           = 200' A string value (Parameter object only).
    adLongVarChar       = 201' A long string value.
    adVarWChar          = 202' A null-terminated Unicode character string.
    adLongVarWChar      = 203' A long null-terminated Unicode string value.
    adVarBinary         = 204' A binary value (Parameter object only).
    adLongVarBinary     = 205' A long binary value.
    'AdArray             = 0x2000' A flag value combined with another data type constant. Indicates an array of that other data type.
' ADODB LockType Constants
    ' Unspecified type of lock. Clones inherits lock type from the original RecordSet.
    adLockUnspecified     = -1
    ' Read-only records
    adLockReadOnly        = 1
    ' Pessimistic locking, record by record. The provider lock records immediately after editing
    adLockPessimistic     = 2
    ' Optimistic locking, record by record. The provider lock records only when calling update
    adLockOptimistic      = 3
    ' Optimistic batch updates. Required for batch update mode
    adLockBatchOptimistic =	4
' ADODB Cursor Type Constants
    ' Does not specify the type of cursor.
    dOpenUnspecified  = -1
    ' Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a RecordSet.
    adOpenForwardOnly = 0
    ' Uses a keySet cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your RecordSet. Data changes by other users are still visible.
    adOpenKeySet	  = 1
    ' Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the RecordSet are allowed, except for bookmarks, if the provider doesn't support them.
    adOpenDynamic	  = 2
    ' Uses a static cursor. A static copy of a Set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
    adOpenStatic	  = 3
' ADODB Cursor Location Constants
    'OBSOLETE (appears only for backward compatibility). Does not use cursor services
    adUseNone	= 1
    'Default. Uses a server-side cursor
    adUseServer	= 2
    'Uses a client-side cursor supplied by a local cursor library. For backward compatibility, the synonym adUseClientBatch is also supported
    adUseClient	= 3
%>