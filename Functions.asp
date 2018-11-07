<%
' Type conversion/checking functions
    ' Converts $Value to string if its not Null or Empty.
    ' Else convert it to Null.
    '
    ' @param {mixed} Value
    ' @return string|null
    Function AsString( Value )
        if IsNull(Value) or IsEmpty(Value) then
            AsString = Null
        else
            AsString = Trim(CStr(Value))
        end if
    End Function
    ' Checks if $Value is a string.
    '
    ' @param {mixed} Value
    ' @return bool
    Function IsString( Value )
        IsString = (TypeName(Value) = "String")

        if not IsString then
            IsString = not IsNull(AsString( Value ))
        end if
    End Function

    ' Checks the given date-string is in ISO 8601.
    ' @param {string} s_input
    ' @return bool
    function isIsoDate(s_input)
        dim obj_regex

        isIsoDate = false
        if len(s_input) > 9 then ' basic check before creating RegExp
            set obj_regex = new RegExp
            obj_regex.Pattern = "^\d{4}\-\d{2}\-\d{2}(T\d{2}:\d{2}:\d{2}(Z|\+\d{4}|\-\d{4})?)?$"
            if obj_regex.Test(s_input) then
                on error resume next
                isIsoDate = not IsEmpty(CIsoDate(s_input))
                on error goto 0
            end if
            set obj_regex = nothing
        end if
    end function

    ' Parses the given date-string from ISO 8601.
    ' @param {string} s_input
    ' @return Date
    function CIsoDate(s_input)
        CIsoDate = CDate(replace(Mid(s_input, 1, 19) , "T", " "))
    end function
    ' Parses the given date to an ISO 8601 string.
    ' @param {Date} d_input
    ' @return {sting}
    function CIsoString(d_input)
        ' 2018-08-07T19:40Z
        CIsoString = _
            ZeroFill(Year(d_input), 4) & "-" &_
            ZeroFill(Month(d_input), 2) & "-" &_
            ZeroFill(Day(d_input), 2) &_
        "T" & "-" &_
            ZeroFill(Hour(d_input), 2) & "-" &_
            ZeroFill(Minute(d_input), 2) & "Z"
    end function

    ' Converts $Value to Date if its a valid date value.
    ' Else convert it to Null.
    '
    ' @param {mixed} Value
    ' @return Date|null
    Function AsDate( Value )
        On Error Resume Next
            if IsNull(Value) or IsEmpty(Value) then
                AsDate = Null
            elseif isIsoDate( Value ) then
                AsDate = CIsoDate(Value)
            elseif IsDate(Value) or IsString(Value) then
                AsDate = CDate(Value)
            else
                AsDate = Null
            end if
        On Error Goto 0

        if IsEmpty(AsDate) then
            AsDate = Null
        end if
    End Function

    ' Converts $Value to Bool if its a valid boolean value.
    ' Else convert it to Null.
    '
    ' @param {mixed} Value
    ' @return bool|null
    Function AsBool( Value )
        On Error Resume Next
            if IsNull(Value) or IsEmpty(Value) then
                AsBool = Null
            else
                AsBool = CBool(Value)
            end if
        On Error Goto 0

        if IsEmpty(AsBool) then
            AsBool = Null
        end if
    End Function
    ' Checks if $Value is a boolean.
    '
    ' @param {mixed} Value
    ' @return bool
    Function IsBool( Value )
        IsBool = (TypeName(Value) = "Boolean")

        if not IsBool then
            IsBool = not IsNull(AsBool( Value ))
        end if
    End Function

    ' Converts $Value to Byte if its a valid byte value.
    ' Else convert it to Null.
    '
    ' @param {mixed} Value
    ' @return byte|null
    Function AsByte( Value )
        On Error Resume Next
            if IsNull(Value) or IsEmpty(Value) then
                AsByte = Null
            else
                AsByte = CByte(Value)
            end if
        On Error Goto 0

        if IsEmpty(AsByte) then
            AsByte = Null
        end if
    End Function
    ' Checks if $Value is a byte.
    '
    ' @param {mixed} Value
    ' @return bool
    Function IsByte( Value )
        IsByte = (TypeName(Value) = "Byte")

        if not IsByte then
            IsByte = not IsNull(AsByte( Value ))
        end if
    End Function

    ' Converts $Value to Integer if its a valid int value.
    ' Else convert it to Null.
    '
    ' @param {mixed} Value
    ' @return int|null
    Function AsInt( Value )
        On Error Resume Next
            if IsNull(Value) or IsEmpty(Value) then
                AsInt = Null
            else
                AsInt = CInt(Value)
            end if
        On Error Goto 0

        if IsEmpty(AsInt) then
            AsInt = Null
        end if
    End Function
    ' Checks if $Value is a integer.
    '
    ' @param {mixed} Value
    ' @return bool
    Function IsInt( Value )
        IsInt = (TypeName(Value) = "Integer")

        if not IsInt then
            IsInt = not IsNull(AsInt( Value ))
        end if
    End Function

    ' Converts $Value to Long if its a valid long value.
    ' Else convert it to Null.
    '
    ' @param {mixed} Value
    ' @return long|null
    Function AsLong( Value )
        On Error Resume Next
            if IsNull(Value) or IsEmpty(Value) then
                AsLong = Null
            else
                AsLong = CLng(Value)
            end if
        On Error Goto 0

        if IsEmpty(AsLong) then
            AsLong = Null
        end if
    End Function
    ' Checks if $Value is a long.
    '
    ' @param {mixed} Value
    ' @return bool
    Function IsLong( Value )
        IsLong = (TypeName(Value) = "Long")

        if not IsLong then
            IsLong = not IsNull(AsLong( Value ))
        end if
    End Function

    ' Converts $Value to Double if its a valid double value.
    ' Else convert it to Null.
    '
    ' @param {mixed} Value
    ' @return double|null
    Function AsDouble( Value )
        On Error Resume Next
            if IsNull(Value) or IsEmpty(Value) then
                AsDouble = Null
            else
                AsDouble = CDbl(Value)
            end if
        On Error Goto 0

        if IsEmpty(AsDouble) then
            AsDouble = Null
        end if
    End Function
    ' Checks if $Value is a double.
    '
    ' @param {mixed} Value
    ' @return bool
    Function IsDouble( Value )
        IsDouble = (TypeName(Value) = "Double")

        if not IsDouble then
            IsDouble = not IsNull(AsDouble( Value ))
        end if
    End Function

    ' Converts $Value to Single if its a valid single value.
    ' Else convert it to Null.
    '
    ' @param {mixed} Value
    ' @return single|null
    Function AsSingle( Value )
        On Error Resume Next
            if IsNull(Value) or IsEmpty(Value) then
                AsSingle = Null
            else
                AsSingle = CSng(Value)
            end if
        On Error Goto 0

        if IsEmpty(AsSingle) then
            AsSingle = Null
        end if
    End Function
    ' Checks if $Value is a single.
    '
    ' @param {mixed} Value
    ' @return bool
    Function IsSingle( Value )
        IsSingle = (TypeName(Value) = "Single")

        if not IsSingle then
            IsSingle = not IsNull(AsSingle( Value ))
        end if
    End Function

    ' Converts $Value to Currency if its a valid currency value.
    ' Else convert it to Null.
    '
    ' @param {mixed} Value
    ' @return currency|null
    Function AsCurrency( Value )
        On Error Resume Next
            if IsNull(Value) or IsEmpty(Value) then
                AsCurrency = Null
            else
                AsCurrency = CCur(Value)
            end if
        On Error Goto 0

        if IsEmpty(AsCurrency) then
            AsCurrency = Null
        end if
    End Function
    ' Checks if $Value is a currency.
    '
    ' @param {mixed} Value
    ' @return bool
    Function IsCurrency( Value )
        IsCurrency = (TypeName(Value) = "Currency")

        if not IsCurrency then
            IsCurrency = not IsNull(AsCurrency( Value ))
        end if
    End Function
    ' Checks if $Value has no data.
    ' Empty, Nothing and Null are values not considered as data.
    '
    ' @param {mixed} Value
    ' @return {bool}
    Function IsVoid( Value )
        IsVoid = IsEmpty(Value)

        if not IsVoid then
            IsVoid = IsNull(Value)

            if isObject(Value) and not IsVoid then
                IsVoid = Value is Nothing
            end if
        end if
    End Function
' Array functions
    ' Creates an empty array with $Size elements allocated.
    ' @param {int} Size
    ' @return Array
    Function EmptyArray( Size )
        Dim Result()
        ReDim Result(Size)

        EmptyArray = Result
    End Function
    ' Creates an empty Dictionary.
    ' @return {Scripting.Dictionary}
    Function Dictionary( )
        Set Dictionary = CreateObject("Scripting.Dictionary")
    End Function
' SQL inspired functions
    ' Functional-equivalent of ternary operator.
    '
    ' @param {bool} Condition
    ' @param {mixed} ValidReturn
    ' @param {mixed} InvalidReturn
    ' @return {mixed}
    Function IIF(Condition, ValidReturn, InvalidReturn)
        if (Condition) then
            set_ IIF, ValidReturn
        else
            set_ IIF, InvalidReturn
        end if
    End Function
' Utilitary functions
    function ZeroFill(Value, Length)
        Dim ValueLength
        Dim FillLength
         if IsVoid(Value) then
            Value = ""
        end if
         ValueLength = LEN(CStr(Value))
        FillLength = Length - ValueLength
         if FillLength > 0 then
            ZeroFill = Replace(Space(FillLength), " ", "0") & Value
        else
            ZeroFill = Value
        end if
    end function
%>