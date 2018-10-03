<%
' Encapsulates the SQL statements logic in a much simpler to use class.
'
' Allows arbitrary clause order when through the public *_Clause functions.
' Outputs ready-to-use ADODB.Command objects by the Build_* functions,
' automatically handling parameters values and types.
Class DB_Statement
    ' Internal interface
        ' @var {Dictionary} Clauses of the current statement
        Private Clauses
        ' Gets a clause on current statement, creating it if that's needed.
        '
        ' @return {Dictionary}
        Private Property Get Clause( Name )
            if not Clauses.Exists(Name) then
                Set Clauses(Name) = Dictionary()
            end if

            Set Clause = Clauses(Name)
        End Property
        ' Set Value(s) on clause $Name, filling with Filler when $Value has
        ' only keys.
        '
        ' @param {string} Name clause name
        ' @param {mixed} Filler default value to fill lists of keys, must be a
        ' scalar type
        ' @param {string|Array|Dictionary} Value
        Private Property Let ClauseValue( Name, Field, Value)
            Clause(Name)(Field) = Value
        End Property
        ' Sets Value(s) on clause $Name, filling with Filler when $Values has
        ' only keys.
        '
        ' @param {string} Name clause name
        ' @param {mixed} Filler default value to fill lists of keys, must be a
        ' scalar type
        ' @param {string|Array|Dictionary} Values
        Private Property Let ClauseValues( Name, Filler, Values)
            Dim Key
            Dim CurrentClause : Set CurrentClause = Clause(Name)

            Select Case TypeName(Values)
                Case "Variant()" ' Array
                    For Each Key in Values
                        ' No Alias
                        CurrentClause(Key) = Filler
                    Next
                Case "Dictionary"
                    For Each Key in Values
                        ' Key: Alias"
                        CurrentClause(Key) = Values(Key)
                    Next
                Case "String"
                    ' No Alias
                    CurrentClause(Values) = Filler
            End Select
        End Property
        ' Appends a new ADODB.Parameter to $Command, using given $Name and
        ' $Value.
        '
        ' @param {ADODB.Command} Command
        ' @param {string} Name
        ' @param {mixed} Value
        Private Sub AppendParameter(Byref Command, Name, Value)
            Dim Parameter : Set Parameter = Command.CreateParameter(Name, , 1, , Value)
            Call Command.Parameters.Append(FixType(Parameter))
        End Sub
        ' Adds $Conditions to $Command.
        '
        ' @param {ADODB.Command} Command
        ' @param {Dictionart<int, DB_Condition>} Conditions
        Private Sub ApplyConditions( ByRef Command, ByRef Conditions )
            Dim Condition
            Dim Index
            Dim Key

            Dim First : First = true

            For Each Key in Conditions
                Set Condition = Conditions(Key)

                ' Skips operator in first condition
                if First then
                    First = False
                else
                    Command.CommandText = _
                        Command.CommandText & " " & Condition.Operator & " "
                end if

                Command.CommandText = Command.CommandText & Condition.Field & " "

                Select Case TypeName(Condition.Values)
                    Case "Empty"
                        ' Do nothing
                    Case "Null", "Nothing"
                        Command.CommandText = Command.CommandText & " IS NULL"
                    Case "Variant()"
                        Command.CommandText = Command.CommandText & " IN("
                        For Index = 0 To Ubound(Condition.Values) Step 1
                            if IsNull(Condition.Values(Index)) then
                                Command.CommandText = Command.CommandText & "null,"
                            elseif not IsEmpty(Condition.Values(Index)) then
                                Command.CommandText = Command.CommandText & "?,"

                                Set Parameter = _
                                Command.CreateParameter(Condition.Field & "_" & Index, , 1, , Condition.Values(Index))

                                Call Command.Parameters.Append(FixType(Parameter))
                            end if
                        Next
                        Command.CommandText = _
                            LEFT(Command.CommandText, LEN(Command.CommandText) - 1) & ")"
                    Case Else
                        Command.CommandText = Command.CommandText & " = ?"
                        
                        Set Parameter = Command.CreateParameter(Condition.Field, , 1, , Condition.Values)
                        Call Command.Parameters.Append(FixType(Parameter))
                End Select
            Next
        End Sub
        ' Adds JOIN clauses to $Command.
        '
        ' @param {ADODB.Command} Command
        Private Sub ApplyJoins( ByRef Command )
            if Clauses.Exists("{{JOIN}}") then
                Dim Key
                Dim JoinClause

                Dim CurrentClause : Set CurrentClause = Clauses("{{JOIN}}")

                For Each Key in CurrentClause
                    set JoinClause = CurrentClause(Key)
                    Command.CommandText = Command.CommandText & " " &_
                        JoinClause.JoinType & " JOIN " & _
                        JoinClause.Table & " ON "

                    Call ApplyConditions(Command, JoinClause.Conditions)
                Next
            end if
        End Sub
        ' Adds WHERE clause to $Command.
        '
        ' @param {ADODB.Command} Command
        Private Sub ApplyWhere( ByRef Command )
            if Clauses.Exists("{{WHERE}}") then
                Command.CommandText = Command.CommandText & " WHERE "

                Call ApplyConditions(Command, Clauses("{{WHERE}}"))
            end if
        End Sub
        ' Checks the type of $Parameter.Value with VBScript's TypeName
        ' function, assigning the right ADODB type in $Parameter.
        ' In case of String types, also assigns $Parameter.Size.
        '
        ' This function is ready to work with Entities by reflection.
        '
        ' @param {ADODB.Parameter} Parameter
        ' @return {ADODB.Parameter} The same input parameter (syntax sugar)
        Private Function FixType(ByRef Parameter)
            Dim DetectType
    		Dim Length

            Select Case TypeName(Parameter.Value)
                Case "String":
                    Length = LEN(Parameter.Value)
                    if Length = 0 then
                        Length = 1
                    end if

                    Parameter.Size = Length

                    if UseUTF then
                        Parameter.Type = adLongVarWChar
                    else
                        Parameter.Type = adLongVarChar
                    end if
                Case "Integer"
                    Parameter.Type = adInteger
                Case "Long"
                    Parameter.Type = adBigInt
                Case "Byte"
                    Parameter.Type = adSmallInt
                Case "Single"
                    Parameter.Type = adSingle
                Case "Double"
                    Parameter.Type = adDouble
                Case "Currency"
                    Parameter.Type = adCurrency
                Case "Date"
                    Parameter.Type = adDBTimeStamp
                Case "Boolean"
                    Parameter.Type = adBoolean
            End Select

            Set FixType = Parameter
        End Function



    ' Public interface
        ' If automatic type detection will mark VBScript strings as UTF8.
        '
        ' @var {bool}
        Public UseUTF
        ' If built SQL Statements are Prepared Statements.
        '
        ' @var {bool}
        Public PrepareCommands
        ' Removes all caluses from this statement.
        '
        ' @return {self}
        Public Function Clear( )
            Call Clauses.RemoveAll()

            Set Clear = Me
        End Function



    ' Pre-constructor
        Sub Class_Initialize( )
            Set Clauses = Dictionary()
            UseUTF = False
            PrepareCommands = False
        End Sub



    ' Destructor
        Sub Class_Terminate( )
            Call Clauses.RemoveAll()
            Set Clauses = Nothing
        End Sub



    ' Table-related clauses
        ' Set the table for Select/Delete statements.
        '
        ' @param {string} Table
        ' @return {self}
        Public Function From_Clause(Table)
            if not IsString(Table) then
                Call Err.Raise( _
                    13, _
                    "Statement.From_Clause", _
                    "Table name must be a string" _
                )
            end if

            Clauses("{{FROM}}") = Table

            Set From_Clause = Me
        End Function
        ' Set the table for Insert statements.
        '
        ' @param {string} Table
        ' @return {self}
        Public Function Into_Clause(Table)
            if not IsString(Table) then
                Call Err.Raise( _
                    13, _
                    "Statement.Into_Clause", _
                    "Table name must be a string" _
                )
            end if

            Clauses("{{INTO}}") = Table

            Set Into_Clause = Me
        End Function
        ' Set the table for Update statements.
        '
        ' @param {string} Table
        ' @return {self}
        Public Function Update_Clause(Table)
            if not IsString(Table) then
                Call Err.Raise( _
                    13, _
                    "Statement.Update_Clause", _
                    "Table name must be a string" _
                )
            end if

            Clauses("{{UPDATE}}") = Table

            Set Update_Clause = Me
        End Function



    ' Join-related clauses
        ' Adds a JOIN clause of $JoinType for $Table, with $Conditions, using
        ' $Operator for them.
        '
        ' @param {string} JoinType
        ' @param {string} Table
        ' @param {string|array|Dictionary|DB_Condition} Conditions
        ' @param {string} Operator
        ' @return {self}
        Public Function Join_Clause(JoinType, Table, Conditions, Operator)
            if IsEmpty(Operator) or IsNull(Operator) then
                Operator = "AND"
            elseif IsString(Operator) then
                Operator = UCase(Operator)
            else
                Call Err.Raise( _
                    13, _
                    "Statement.Join_Clause", _
                    "Condition operator must be a string" _
                )
            end if

            if IsEmpty(JoinType) or IsNull(JoinType) then
                JoinType = "INNER"
            elseif IsString(JoinType) then
                JoinType = UCase(JoinType)
            else
                Call Err.Raise( _
                    13, _
                    "Statement.Join_Clause", _
                    "Join type must be a string" _
                )
            end if

            if not IsString(Table) then
                Call Err.Raise( _
                    13, _
                    "Statement.Join_Clause", _
                    "Table must be a string" _
                )
            end if

            Dim Condition
            Dim Key
            Dim CurrentClause : Set CurrentClause = Clause("{{JOIN}}")
            Dim JoinClause    : Set JoinClause    = (new DB_JoinClause)(JoinType, Table)

            Select Case TypeName(Conditions)
                Case "Variant()"
                    For Each Key in Conditions
                        if TypeName(Key) = "DB_Condition" then
                            Set JoinClause.Conditions(JoinClause.Conditions.Count) = Key
                        else
                            Set JoinClause.Conditions(JoinClause.Conditions.Count) = _
                                (new DB_Condition)(Operator, Key, Empty)
                        end if
                    Next
                Case "Dictionary"
                    For Each Key in Conditions
                        if TypeName(Conditions(Key)) = "DB_Condition" then
                            Set JoinClause.Conditions(JoinClause.Conditions.Count) = _
                                Conditions(Key)
                        else
                            Set JoinClause.Conditions(JoinClause.Conditions.Count) = _
                                (new DB_Condition)(Operator, Key, Conditions(Key))
                        end if
                    Next
                Case "DB_Condition"
                    Set JoinClause.Conditions(JoinClause.Conditions.Count) = _
                        Conditions
                Case "String"
                    ' Avoiding function call overhead
                    Set JoinClause.Conditions(JoinClause.Conditions.Count) = _
                        (new DB_Condition)(Operator, Conditions, Empty)

            End Select

            Set CurrentClause(CurrentClause.Count) = JoinClause

            Set Join_Clause = Me
        End Function



    ' Field-related clauses
        ' Stores the given $Fields in the fieldlist for Select statements
        ' To provide alisases, pass a dictionary in $Fields.
        '
        ' @param {string|Array|Dictionary} Fields
        ' @return {self}
        Public Function Select_Clause(Fields)
            ClauseValues("{{SELECT}}", Empty) = Fields

            Set Select_Clause = Me
        End Function
        ' Stores the given $Fields in the field/value list for Update
        ' statements.
        ' To provide values to the fields pass a Dictionary to $Fields,
        ' or they will be set to null.
        '
        ' @param {string|Array|Dictionary} Fields
        ' @return {self}
        Public Function Set_Clause(Fields)
            ClauseValues("{{SET}}", Empty) = Fields

            Set Set_Clause = Me
        End Function
        ' Stores $Field and $Value in the field/value list for Update
        ' statements.
        '
        ' @param {string} Field
        ' @param {mixed} Value
        ' @return {self}
        Public Function Set_Field(Field, Value)
            ClauseValue("{{SET}}", Field) = Value

            Set Set_Field = Me
        End Function
        ' Stores the given $Fields in the field/value list for Insert
        ' statements.
        ' To provide values to the fields pass a Dictionary to $Fields,
        ' or they will be set to null.
        '
        ' @param {string|Array|Dictionary} Fields
        ' @return {self}
        Public Function Insert_Clause(Fields)
            ClauseValues("{{INSERT}}", Empty) = Fields

            Set Insert_Clause = Me
        End Function
        ' Stores $Field and $Value in the field/value list for Insert
        ' statements.
        '
        ' @param {string} Field
        ' @param {mixed} Value
        ' @return {self}
        Public Function Insert_Field(Field, Value)
            ClauseValue("{{INSERT}}", Field) = Value

            Set Insert_Field = Me
        End Function



    ' Condition-related clauses
        ' Adds conditions on WHERE clause, using $Operator for them.
        '
        ' @param {string|array|Dictionary|DB_Condition} Conditions
        ' @param {string} Operator
        ' @return {self}
        Public Function Where_Clause(Conditions, Operator)
            if IsEmpty(Operator) or IsNull(Operator) then
                Operator = "AND"
            elseif IsString(Operator) then
                Operator = UCase(Operator)
            else
                Call Err.Raise( _
                    13, _
                    "Statement.Where_Clause", _
                    "Condition operator must be a string" _
                )
            end if

            Dim CurrentClause : Set CurrentClause = Clause("{{WHERE}}")
            Dim Condition
            Dim Key

            Select Case TypeName(Conditions)
                Case "Variant()"
                    For Each Key in Conditions
                        if TypeName(Key) = "DB_Condition" then
                            Set CurrentClause(CurrentClause.Count) = Key
                        else
                            Set CurrentClause(CurrentClause.Count) = _
                                (new DB_Condition)(Operator, Key, Empty)
                        end if
                    Next
                Case "Dictionary"
                    For Each Key in Conditions
                        if TypeName(Conditions(Key)) = "DB_Condition" then
                            Set CurrentClause(CurrentClause.Count) = Conditions(Key)
                        else
                            Set CurrentClause(CurrentClause.Count) = _
                                (new DB_Condition)(Operator, Key, Conditions(Key))
                        end if
                    Next
                Case "DB_Condition"
                    Set CurrentClause(CurrentClause.Count) = Conditions
                Case "String"
                    ' Avoiding overhed of calling Where_In()
                    Set CurrentClause(CurrentClause.Count) = _
                        (new DB_Condition)(Operator, Conditions, Empty)
            End Select

            Set Where_Clause = Me
        End Function
        ' Adds a condition on WHERE clause to check if $Field is equal/in
        ' $Values, using $Operator for this condition.
        '
        ' @param {string} Field
        ' @param {string|array} Values
        ' @param {string} Operator
        ' @return {self}
        Public Function Where_In(Field, Values, Operator)
            if IsEmpty(Operator) or IsNull(Operator) then
                Operator = "AND"
            elseif IsString(Operator) then
                Operator = UCase(Operator)
            else
                Call Err.Raise( _
                    13, _
                    "Statement.Where_In", _
                    "Condition operator must be a string" _
                )
            end if
            if not IsString(Field) then
                Call Err.Raise( _
                    13, _
                    "Statement.Where_In", _
                    "Field name must be a string" _
                )
            end if

            Dim CurrentClause : Set CurrentClause = Clause("{{WHERE}}")

            Set CurrentClause(CurrentClause.Count) = _
                (new DB_Condition)(Operator, Field, Values)

            Set Where_In = Me
        End Function






    ' Order-related clauses
        ' Adds a field to ORDER BY clause.
        '
        ' @param {string|array<string>} Fields
        ' @param {string} Order
        ' @return {self}
        Public Function Order_Clause( Fields, Order )
            if Not IsEmpty(Order) then
                Order = UCase(Order)
            end if
            ClauseValues("{{ORDER}}", Order) = Fields

            Set Order_Clause = Me
        End Function



    ' Group-related clauses
        ' Adds a field to GROUP BY clause.
        '
        ' @param {string|array<string>} Fields
        ' @return {self}
        Public Function Group_Clause( Fields )
            ClauseValues("{{GROUP}}", Empty) = Fields

            Set Group_Clause = Me
        End Function



    ' SQL Statement Assemble
        ' Assembles the clauses from this statement in an INSERT statement.
        ' After assembling the SQL, this DB_Statemet object is cleared.
        '
        ' @return {ADODB.Command}
        Public Function Build_Insert( )
            if Clauses.Exists("{{INTO}}") then
                if Clauses.Exists("{{INSERT}}") then
                    Dim Command       : Set Command       = CreateObject("ADODB.Command")
                    Dim CurrentClause : Set CurrentClause = Clauses("{{INSERT}}")
                    Dim ValuesString
                    Dim Parameter
                    Dim Key
                    Dim Value

                    Command.CommandText = "INSERT INTO " & Clauses("{{INTO}}") & " ("
                    Command.Prepared = PrepareCommands
                    ValuesString = ") VALUES ("

                    For Each Key in CurrentClause
                        Value = CurrentClause(Key)
                        if not IsEmpty(Value) then
                            Command.CommandText = Command.CommandText & Key & ","

                            if IsNull(Value) then
                                ValuesString = ValuesString & "null,"
                            else
                                ValuesString = ValuesString & "?,"

                                Call AppendParameter(Command, Key, Value)
                            end if
                        end if
                    Next

                    Command.CommandText = _
                        LEFT(Command.CommandText, LEN(Command.CommandText) - 1) & _
                        LEFT(ValuesString, LEN(ValuesString) - 1) & ")"

                    Call Clear()

                    set Build_Insert = Command
                else
                    Call Err.Raise(448, "No data to insert into table, missing Insert_Clause(Fields) call", "DB_Statement.Build_Insert")
                end if
            else
                Call Err.Raise(448, "No table to insert into, missing Into_Clause(Table) call", "DB_Statement.Build_Insert")
            end if
        End Function
        ' Assembles the clauses from this statement in a SELECT statement.
        ' After assembling the SQL, this DB_Statemet object is cleared.
        '
        ' @return {ADODB.Command}
        Public Function Build_Select( )
            if Clauses.Exists("{{FROM}}") then
                Dim CurrentClause
                Dim Parameter
                Dim Key
                Dim Value

                Dim Command : Set Command = CreateObject("ADODB.Command")

                Command.CommandText = "SELECT "
                Command.Prepared = PrepareCommands

                if Clauses.Exists("{{SELECT}}") then
                    Set CurrentClause = Clause("{{SELECT}}")

                    For Each Key in CurrentClause
                        Value = CurrentClause(Key)
                        if IsEmpty(Value) then
                            Command.CommandText = Command.CommandText & Key & ","
                        else
                            Command.CommandText = Command.CommandText & Key & " AS " & Value & ","
                        end if
                    Next

                    Command.CommandText = LEFT(Command.CommandText, LEN(Command.CommandText) - 1)
                else
                    Command.CommandText = Command.CommandText & "*"
                end if

                Command.CommandText = Command.CommandText & " FROM " & Clauses("{{FROM}}")
                Call ApplyJoins(Command)
                Call ApplyWhere(Command)

                if Clauses.Exists("{{GROUP}}") then
                    Set CurrentClause = Clause("{{GROUP}}")
                    Command.CommandText = Command.CommandText & " GROUP BY "

                    For Each Key in CurrentClause
                        Command.CommandText = Command.CommandText & Key & ","
                    Next

                    Command.CommandText = LEFT(Command.CommandText, LEN(Command.CommandText) - 1)
                end if

                if Clauses.Exists("{{ORDER}}") then
                    Set CurrentClause = Clause("{{ORDER}}")
                    Command.CommandText = Command.CommandText & " ORDER BY "

                    For Each Key in CurrentClause
                        if IsEmpty(CurrentClause(Key)) then
                            Command.CommandText = Command.CommandText & Key & ","
                        else
                            Command.CommandText = Command.CommandText & Key & " " & CurrentClause(Key) & ","
                        end if
                    Next

                    Command.CommandText = LEFT(Command.CommandText, LEN(Command.CommandText) - 1)
                end if

                Call Clear()

                Set Build_Select = Command
            else
                Call Err.Raise(448, "No table to select, missing From_Clause(Table) call", "DB_Statement.Build_Select")
            end if
        End Function
        ' Assembles the clauses from this statement in an UPDATE statement.
        ' After assembling the SQL, this DB_Statemet object is cleared.
        '
        ' @return {ADODB.Command}
        Public Function Build_Update( )
            if Clauses.Exists("{{UPDATE}}") then
                if Clauses.Exists("{{SET}}") then
                    Dim Command       : Set Command       = CreateObject("ADODB.Command")
                    Dim CurrentClause : Set CurrentClause = Clauses("{{SET}}")
                    Dim Parameter
                    Dim Key
                    Dim Value

                    Command.CommandText = "UPDATE " & Clauses("{{UPDATE}}") & " SET "
                    Command.Prepared = PrepareCommands

                    For Each Key in CurrentClause
                        Value = CurrentClause(Key)
                        if not IsEmpty(Value) then
                            if IsNull(Value) then
                                Command.CommandText = Command.CommandText & Key & "=null,"
                            else
                                Command.CommandText = Command.CommandText & Key & "=?,"

                                Call AppendParameter(Command, Key, Value)
                            end if
                        end if
                    Next

                    Command.CommandText = LEFT(Command.CommandText, LEN(Command.CommandText) - 1)

                    Call ApplyWhere(Command)

                    Call Clear()

                    Set Build_Update = Command
                else
                    Call Err.Raise(448, "No data to update table, missing Set_Clause(Fields) call", "DB_Statement.Build_Update")
                end if
            else
                Call Err.Raise(448, "No table to update, missing Update_Clause(Table) call", "DB_Statement.Build_Update")
            end if
        End Function
        ' Assembles the clauses from this statement in an DELETE statement.
        ' After assembling the SQL, this DB_Statemet object is cleared.
        '
        ' @return {ADODB.Command}
        Public Function Build_Delete( )
            if Clauses.Exists("{{FROM}}") then
                Dim Command : Set Command = CreateObject("ADODB.Command")

                Command.CommandText = "DELETE FROM " & Clauses("{{FROM}}")
                Command.Prepared = PrepareCommands

                Call ApplyWhere(Command)

                Call Clear()

                Set Build_Delete = Command
            else
                Call Err.Raise(448, "No table to delete from, missing From_Clause(Table) call", "DB_Statement.Build_Delete")
            end if
        End Function
End Class
%>