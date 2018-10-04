<!--#include file="ADOConstants.asp"-->
<!--#include file="Functions.asp"-->
<!--#include file="Conditions.asp"-->
<!--#include file="Statement.asp"-->
<%
' Encapsulates database operations in a simple-to-use object.
Class DataBase
    ' Internal Interface
        ' @var {int}
        Private ConnectionCount
        ' @var {DB_Statement}
        Private CurrentStatement
        ' @var {Dictionary(string, Entity)}
        Private LoadedEntities
        
        Private Function FixEntityField(DataType, Value)
            If IsNull(Value) Then
                FixEntityField = Null
            Else
                Dim Conversion

                Select Case DataType
                    Case adBoolean
                        FixEntityField = CBool(Value)
                    Case adSmallInt, adTinyInt, adUnsignedTinyInt
                        FixEntityField = CByte(Value)
                    Case adCurrency
                        FixEntityField = CCur(Value)
                    Case adDouble
                        FixEntityField = CDbl(Value)
                    Case adDate, adFileTime, adDBDate, adDBTime, adDBTimeStamp
                        if MySQL_Date_Patch then
                            FixEntityField = CStr(Value)
                        else
                            FixEntityField = CDate(Value)
                        end if
                    Case adInteger, adUnsignedSmallInt
                        FixEntityField = CInt(Value)
                    Case adBigInt, adUnsignedInt, adError
                        FixEntityField = CLng(Value)
                    Case adSingle
                        FixEntityField = CSng(Value)
                    Case adBSTR, adLongVarChar, adLongVarWChar, _
                    adDecimal, adUnsignedBigInt, adNumeric, _
                    adGUID, adChar, adWChar, adVarChar, adVarWChar
                        FixEntityField = CStr(Value)
                    Case Else
                        return FixEntityField, Value
                End Select
            End if
        End Function



    ' Public interface
        ' @var {ADODB.Connection}
        Public Connection
        ' If the data-type correction for MySQL must be used.
        '
        ' @var {bool}
        Public MySQL_Date_Patch
        ' If FowardOnly recordsets must be used en queries.
        ' @var {bool}
        Public UseFowardOnly

        ' Removes all caluses from this statement.
        '
        ' @return {self}
        Public Function Clear( )
            Call CurrentStatement.Clear()

            Set Clear = Me
        End Function
        ' Creates a Recordset disconnected from database, allowing to use DB
        ' data without an active connection to it.
        ' Uses Foward Only cursor type to maxime performance.
        '
        ' @return {ADODB.Recordset}
        Public Function ForwardOnlyRecordset( )
            Dim Recordset : Set Recordset = CreateObject("ADODB.Recordset")

            Recordset.CursorLocation = adUseClient
            Recordset.CursorType     = adOpenForwardOnly
            Recordset.LockType       = adLockReadOnly

            Set ForwardOnlyRecordset = Recordset
        End Function
        ' Creates a Recordset disconnected from database, allowing to use DB
        ' data without an active connection to it.
        ' Uses static cursor type to allow moving in any direction.
        '
        ' @return {ADODB.Recordset}
        Public Function StaticRecordset( )
            Dim Recordset : Set Recordset = CreateObject("ADODB.Recordset")

            Recordset.CursorLocation = adUseClient
            Recordset.CursorType     = adOpenStatic
            Recordset.LockType       = adLockReadOnly

            Set StaticRecordset = Recordset
        End Function
        ' Recovers all the fields of given $Entity.
        '
        ' @param {object} Entity
        ' @return {Dictionary}
        Public Function EntityFields( Entity )
            Dim ClassFields
            Dim Field
            Dim Fields
            Dim Value

            Set Fields = Dictionary()
            Set ClassFields = Class_Loader.Members(Entity.Self)
            For Each Field in ClassFields
                set_ Value, Entity(Field)
                If not IsEmpty(Value) Then
                    Fields.Add Field, FixEntityField(ClassFields(Field), Value)
                End if
            Next

            Set EntityFields = Fields
        End Function
        ' Recovers the fields of given $Entity registered as keys.
        '
        ' @param {object} Entity
        ' @return {Dictionary}
        Public Function EntityKeys( Entity )
            Dim ClassFields
            Dim Field
            Dim Fields
            Dim Value

            Set Fields = Dictionary()
            Set ClassFields = Class_Loader.Members(Entity.Self)
            if Entity.Queryable then
                For Each Field in ClassFields
                    set_ Value, Entity(Field)
                    If not IsEmpty(Value) Then
                        Fields.Add Field, FixEntityField(ClassFields(Field), Value)
                    End if
                Next
            elseif ClassFields.Count > 0 then
                Field = Entity.KeyField
                set_ Value, Entity(Field)
                If not IsEmpty(Value) Then
                    Fields.Add Field, FixEntityField(ClassFields(Field), Value)
                End if
            end if

            Set EntityKeys = Fields
        End Function
        ' Converts $Recordset to an array of $Entity objects.
        '
        ' @param {object} Entity
        ' @param {ADODB.Recordset} Recordset
        ' @return {array}
        Public Function ParseEntities( ByRef Entity, ByRef Recordset )
            Dim Current
            Dim Loaded
            Dim Result

            Dim EntityClass : Set EntityClass = Entity.Self

            if not Recordset.EOF then
                Dim Append
                Dim CurrentIndex
                Dim Duplicate
                Dim Index
                Dim LoadIndex
                Dim Size
                Dim Key

                ' Initializing Entity cache on demand
                if IsEmpty(LoadedEntities) then
                    Set LoadedEntities = Dictionary()
                end if

                Append = LoadedEntities.Exists(EntityClass.Name)
                if Append then
                    Loaded = LoadedEntities(EntityClass.Name)
                    LoadIndex = UBound(Loaded)
                    ' RecordCount already adds 1 to the sum
                    Size = Recordset.RecordCount + LoadIndex
                    ReDim Preserve Loaded(Size)

                    Size = Recordset.RecordCount -1
                else
                    Loaded = Array()
                    LoadIndex = -1
                    
                    ' Size count begins in 0
                    Size = Recordset.RecordCount -1
                    ReDim Loaded(Size)
                end if
                Index = 0
                Result = Array()
                ReDim Preserve Result(Size)

                ' Key for duplicate check
                Key = Entity.KeyField
                Set ClassFields = Class_Loader.Members(EntityClass)

                Do Until Recordset.EOF
                    Duplicate = false
                    If Append Then
                        set_ Value, FixEntityField( _
                            EntityClass.Fields(Key), Recordset(Key).Value _
                        )
                        For Each Current in Loaded
                            If IsObject(Current) Then
                                if Current(Key) = Value then
                                    Duplicate = true
                                    Exit For
                                end if
                            Else
                                Exit For
                            End if
                        Next
                    End if

                    If Index = 0 Then
                        set Current = Entity.ToNonQueryable()
                    Else
                        set Current = EntityClass.GetInstance()
                    End if

                    For Each Field in ClassFields
                        Current(Field) = FixEntityField( _
                            ClassFields(Field), Recordset(Field).Value _
                        )
                    Next

                    set Result(Index) = Current
                    Index = Index + 1

                    if not (Append and Duplicate) then
                        LoadIndex = LoadIndex + 1
                        set Loaded(LoadIndex) = Current
                    end if
                    
                    Call Recordset.MoveNext()
                Loop
                Recordset.Close()
                Set Recordset = Nothing
                ' Fiting to content
                ReDim Preserve Loaded(LoadIndex)
                LoadedEntities(EntityClass.Name) = Loaded
            else
                Result = Array()
            end if

            ParseEntities = Result
        End Function



    ' Pre-constructor
        Sub Class_Initialize( )
            Set Connection   = Server.CreateObject("ADODB.Connection")
            ConnectionCount  = 0
            MySQL_Date_Patch = false
            UseFowardOnly    = false

            set CurrentStatement = new DB_Statement
        End Sub



    ' Destructor
        Sub Class_Terminate( )
            Set Connection = Nothing
            Set CurrentStatement = Nothing

            if not IsEmpty(LoadedEntities) then
                Call LoadedEntities.RemoveAll()

                Set LoadedEntities = Nothing
            end if
        End Sub



    ' Connection-related functions
        ' Connects to the database (if not already connected) and increment
        ' the connection counter.
        '
        ' @return {self}
        Public Function Connect( )
            if ConnectionCount = 0 then
                Call Connection.Open()
            end if
            ConnectionCount = ConnectionCount + 1

            Set Connect = Me
        End Function
        ' Disconects from the database (if connected) and decrement the 
        'connection counter.
        '
        ' @return {self}
        Public Function Disconnect( )
            if ConnectionCount > 0 then
                if ConnectionCount = 1 then
                    Call Connection.Close()
                    
                    if not IsVoid(LoadedEntities) then
                        Call LoadedEntities.RemoveAll()
                    
                        Set LoadedEntities = Nothing
                    end if
                end if
                ConnectionCount = ConnectionCount - 1
            end if

            Set Disconnect = Me
        End Function
    


    ' Table-related clauses
        ' Set the table for Select/Delete statements.
        '
        ' @param {string} Table
        ' @return {self}
        Public Function From_Clause(Table)
            Call CurrentStatement.From_Clause(Table)

            Set From_Clause = Me
        End Function
        ' Set the table for Insert statements.
        '
        ' @param {string} Table
        ' @return {self}
        Public Function Into_Clause(Table)
            Call CurrentStatement.Into_Clause(Table)

            Set Into_Clause = Me
        End Function
        ' Set the table for Update statements.
        '
        ' @param {string} Table
        ' @return {self}
        Public Function Update_Clause(Table)
            Call CurrentStatement.Update_Clause(Table)

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
            Call CurrentStatement.Join_Clause(JoinType, Table, Conditions, Operator)

            Set Join_Clause = Me
        End Function
        Public Function Inner_Join(Table, Conditions, Operator)
            Set Inner_Join = Join_Clause("INNER", Table, Conditions, Operator)
        End Function
        Public Function Left_Join(Table, Conditions, Operator)
            Set Left_Join = Join_Clause("LEFT", Table, Conditions, Operator)
        End Function
        Public Function Right_Join(Table, Conditions, Operator)
            Set Right_Join = Join_Clause("RIGHT", Table, Conditions, Operator)
        End Function
        Public Function Outer_Join(Table, Conditions, Operator)
            Set Outer_Join = Join_Clause("FULL OUTER", Table, Conditions, Operator)
        End Function



    ' Field-related clauses
        ' Stores the given $Fields in the fieldlist for Select statements
        ' To provide alisases, pass a dictionary in $Fields.
        '
        ' @param {string|Array|Dictionary} Fields
        ' @return {self}
        Public Function Select_Clause(Fields)
            Call CurrentStatement.Select_Clause(Fields)

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
            Call CurrentStatement.Set_Clause(Fields)

            Set Set_Clause = Me
        End Function
        ' Stores $Field and $Value in the field/value list for Update
        ' statements.
        '
        ' @param {string} Field
        ' @param {mixed} Value
        ' @return {self}
        Public Function Set_Field(Field, Value)
            Call CurrentStatement.Set_Field(Field, Value)

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
            Call CurrentStatement.Insert_Clause(Fields)

            Set Insert_Clause = Me
        End Function
        ' Stores $Field and $Value in the field/value list for Insert
        ' statements.
        '
        ' @param {string} Field
        ' @param {mixed} Value
        ' @return {self}
        Public Function Insert_Field(Field, Value)
            Call CurrentStatement.Insert_Field(Field, Value)

            Set Insert_Field = Me
        End Function



    ' Condition-related clauses
        ' Adds conditions on WHERE clause, using $Operator for them.
        '
        ' @param {string|array|Dictionary|DB_Condition} Conditions
        ' @param {string} Operator
        ' @return {self}
        Public Function Where_Clause(Conditions, Operator)
            Call CurrentStatement.Where_Clause(Conditions, Operator)

            Set Where_Clause = Me
        End Function
        ' Adds conditions on WHERE clause based on given $Entity.
        '
        ' @param {object} Entity
        ' @return {self}
        Public Function Where_Entity( Entity )
            Dim ValidEntity : ValidEntity = IsObject(Entity)
            if ValidEntity then
                ValidEntity = not Entity is Nothing
            end if
            
            if not ValidEntity then
                Call Err.Raise( _
                    13, _
                    "Statement.Where_Entity", _
                    "Entity must be an object" _
                )
            end if

            Dim Conditions : Set Conditions = EntityKeys(Entity)

            if Conditions.Count > 0 then
                Where_Clause Conditions, "AND"
            end if

            Set Conditions = Nothing

            Set Where_Entity = Me
        End Function
        ' Adds a condition on WHERE clause to check if $Field is equal/in
        ' $Values, using $Operator for this condition.
        '
        ' @param {string} Field
        ' @param {string|array} Values
        ' @param {string} Operator
        ' @return {self}
        Public Function Where_In(Field, Values, Operator)
            Call CurrentStatement.Where_In(Field, Values, Operator)

            Set Where_In = Me
        End Function



    ' Order-related clauses
        ' Adds a field to ORDER BY clause.
        '
        ' @param {string|array<string>} Fields
        ' @param {string} Order
        ' @return {self}
        Public Function Order_Clause( Fields, Order )
            Call CurrentStatement.Order_Clause(Fields, Order)

            Set Order_Clause = Me
        End Function



    ' Group-related clauses
        ' Adds a field to GROUP BY clause.
        '
        ' @param {string|array<string>} Fields
        ' @return {self}
        Public Function Group_Clause( Fields )
            Call CurrentStatement.Group_Clause(Fields)

            Set Group_Clause = Me
        End Function



    ' SQL Statement Assemble
        ' Assembles the clauses from this statement in an INSERT statement.
        '
        ' @return {ADODB.Command}
        Public Function Build_Insert( )
            Set Build_Insert = CurrentStatement.Build_Insert()
        End Function
        ' Assembles the clauses from this statement in a SELECT statement.
        '
        ' @return {ADODB.Command}
        Public Function Build_Select( )
            Set Build_Select = CurrentStatement.Build_Select()
        End Function
        ' Assembles the clauses from this statement in an UPDATE statement.
        '
        ' @return {ADODB.Command}
        Public Function Build_Update( )
            Set Build_Update = CurrentStatement.Build_Update()
        End Function
        ' Assembles the clauses from this statement in an DELETE statement.
        '
        ' @return {ADODB.Command}
        Public Function Build_Delete( )
            Set Build_Delete = CurrentStatement.Build_Delete()
        End Function



    ' SQL Statement Execution
        ' Assembles the clauses from this statement in an INSERT statement.
        '
        ' @return {int}
        Public Function Run_Insert( )
            Dim Affected
            Dim Command
            Dim NewConnection : NewConnection = (ConnectionCount = 0)

            Set Command = CurrentStatement.Build_Insert()

            if NewConnection then
                Call Connect()
            end if

            ' Setting Command's Connection
            'Response.Write Command.CommandText & vbcrlf
            'Response.end
            Set Command.ActiveConnection = Connection
            ' Executing statement
            Call Command.Execute(Affected)
            ' Releasing memory
            Set Command.ActiveConnection = Nothing

            if NewConnection then
                Call Disconnect()
            end if

            Run_Insert = Affected
        End Function
        ' Assembles the clauses from this statement in a SELECT statement.
        '
        ' @return {Recordset}
        Public Function Run_Select( )
            Dim Command
            Dim NewConnection : NewConnection = (ConnectionCount = 0)
            Dim Result
            
            if UseFowardOnly then
                Set Result = ForwardOnlyRecordset()
            else
                Set Result = StaticRecordset()
            end if

            Set Command = CurrentStatement.Build_Select()
            if NewConnection then
                Call Connect()
            end if

            ' Setting Command's Connection
            'Response.Write Command.CommandText & vbcrlf
            Set Command.ActiveConnection = Connection
            ' Executing statement
            Call Result.Open(Command)
            ' Releasing memory
            Set Command.ActiveConnection = Nothing
            Set Result.ActiveConnection = Nothing

            if NewConnection then
                Call Disconnect()
            end if

            Set Run_Select = Result
        End Function
        ' Assembles the clauses from this statement in an UPDATE statement.
        '
        ' @return {int}
        Public Function Run_Update( )
            Dim Affected
            Dim Command
            Dim NewConnection : NewConnection = (ConnectionCount = 0)

            Set Command = CurrentStatement.Build_Update()

            if NewConnection then
                Call Connect()
            end if

            'Response.Write Command.CommandText & vbcrlf
            'Response.end
            ' Setting Command's Connection
            Set Command.ActiveConnection = Connection
            ' Executing statement
            Call Command.Execute(Affected)
            ' Releasing memory
            Set Command.ActiveConnection = Nothing

            if NewConnection then
                Call Disconnect()
            end if

            Run_Update = Affected
        End Function
        ' Assembles the clauses from this statement in an DELETE statement.
        '
        ' @return {int}
        Public Function Run_Delete( )
            Dim Affected
            Dim Command : Set Command = CurrentStatement.Build_Delete()
            Dim NewConnection : NewConnection = (ConnectionCount = 0)

            if NewConnection then
                Call Connect()
            end if

            'Response.Write "OE" & vbcrlf
            'Response.Write Command.CommandText & vbcrlf
            'Response.end
            ' Setting Command's Connection
            Set Command.ActiveConnection = Connection
            ' Executing statement
            Call Command.Execute(Affected)
            ' Releasing memory
            Set Command.ActiveConnection = Nothing

            if NewConnection then
                Call Disconnect()
            end if

            Run_Delete = Affected
        End Function




    ' Generic Entity CRUD
        Public Function Create( Entity )
            Call CurrentStatement _
                .Into_Clause(Entity.Self.Field("TableName")) _
                .Insert_Clause(EntityFields(Entity))

            Create = Run_Insert()
        End Function
        ' Read registers from $Entity's table on Databasem using $Entity to
        ' filter records.
        '
        ' @param {object} Entity
        ' @return {array<Entity>}
        Public Function Read( Entity )
            Dim Result
            Dim Recordset
            Dim NewConnection

            Dim PreviousFlag
            PreviousFlag  = UseFowardOnly
            UseFowardOnly = true

            CurrentStatement.From_Clause Entity.Self.Field("TableName")
            Where_Entity Entity

            if not IsEmpty(Entity.Self.Field("OrderField")) then
                CurrentStatement.Order_Clause Entity.Self.Field("OrderField"), Empty
            end if

            NewConnection = (ConnectionCount = 0)
            if NewConnection then
                Call Connect()
            end if

            Result = ParseEntities(Entity, Run_Select())

            if Ubound(Result) <> -1 and TypeName(Entity.Self.Field("Foreign")) = "Dictionary" then
                Read_Foreign Result
            end if

            if NewConnection then
                Call Disconnect()
            end if

            UseFowardOnly = PreviousFlag

            Read = Result
        End Function
        ' Reads all Foreign Entites refered by Primary.
        ' User recursive search of new Entities to load.
        ' 
        ' @param {array<Entity>} Primary
        ' @param {Dictionary<string, string>} Foreign 
        Private Sub Read_Foreign( ByRef Primary )
            Dim PrimaryEntity   : Set PrimaryEntity   = Primary(0)
            Dim PrimaryClass    : Set PrimaryClass    = PrimaryEntity.Self
            Dim ForeignEntities : Set ForeignEntities = PrimaryClass.Field("Foreign")

            if ForeignEntities.Count > 0 then
                Dim AlreadyLoaded
                Dim Duplicate
                Dim Entity
                Dim Field

                Dim Foreign
                Dim ForeignClass
                Dim ForeignEntity
                Dim ForeignFields
                Dim ForeignIndex
                Dim ForeignKey
                Dim ForeignList
                Dim ForeignValue

                Dim PrimaryCount
                Dim PrimaryKey
                Dim PrimaryIndex
                Dim PrimaryField
                Dim PrimaryFields

                PrimaryCount = Ubound(Primary)
                ' Primary Key's 'class'
                PrimaryKey = PrimaryEntity.KeyField
                ' In theory, the Class_Loader will already have the fields in cache
                ' when this function is called from Read()
                Set PrimaryFields = Class_Loader.Members(PrimaryClass)

                ' Each Foreign Key
                For Each PrimaryField in ForeignEntities
                    ' Foreign Key's 'class'
                    Set ForeignClass = Class_Loader(ForeignEntities)
                    ' Foreign Entities fields
                    ' It's useless to update the Class_Loader cache inside the loop.
                    Set ForeignFields = ForeignClass.GetMembers()
                    ' Foreign Key's instance
                    Set ForeignEntity = ForeignClass.GetInstance()

                    ForeignKey = ForeignEntity.KeyField
                    

                    if IsEmpty(PrimaryKey) and IsEmpty(ForeignKey) then
                        Call Err.Raise( _
                            13, _
                            "Database.Read_Foreign", _
                            "Entities " & ForeignClass & " and " & PrimaryClass & " haven't key" _
                        )
                    end if

                    if not (_
                        ForeignFields.Exists(PrimaryKey) _
                        OR PrimaryFields.Exists(ForeignKey) _
                    ) then
                        Call Err.Raise( _
                            13, _
                            "Database.Read_Foreign", _
                            "Entities " & ForeignClass & " and " & PrimaryClass & " doesn't share key" _
                        )
                    end if

                    bReverse = PrimaryFields.Exists(ForeignKey)
                    Field = IIF(bReverse, ForeignKey, PrimaryKey)

                    AlreadyLoaded = false
                    if LoadedEntities.Exists(ForeignClass.Name) then
                        ForeignList = LoadedEntities(ForeignClass.Name)
                        Foreign = EmptyArray(PrimaryCount)

                        ForeignIndex = -1
                        For PrimaryIndex = 0 To PrimaryCount Step 1
                            Duplicate = false
                            ForeignValue = Primary(PrimaryIndex)(Field)
                            For Each Entity in ForeignList
                                if IsObject(Entity) then
                                    if Entity(Field) = ForeignValue then
                                        Duplicate = true
                                        Exit For
                                    end if
                                else
                                    Exit For
                                end if
                            Next
                            if not Duplicate then
                                ForeignIndex = ForeignIndex + 1
                                Foreign(ForeignIndex) = ForeignValue
                            end if
                        Next

                        if ForeignIndex = -1 then
                            AlreadyLoaded = true
                        else
                            ReDim Preserve Foreign(ForeignIndex)
                        end if
                    else
                        Foreign = EmptyArray(PrimaryCount)
                        ForeignIndex = -1
                        For Each Entity in Primary
                            ForeignIndex = ForeignIndex + 1
                            Foreign(ForeignIndex) = Entity(Field)
                        Next
                        ReDim Preserve Foreign(ForeignIndex)
                    end if

                    if AlreadyLoaded then
                        Foreign = LoadedEntities(ForeignClass.Name)
                    elseif UBound(Foreign) <> -1 then
                        Call CurrentStatement _
                            .From_Clause(ForeignClass.Field("TableName") & " AS " & ForeignClass.Name) _
                            .Select_Clause(ForeignClass.Name & ".*")
                        Call CurrentStatement.Where_In( _
                            ForeignClass.Name & "." & Field, Foreign, "AND" _
                        )
                        ' Recyling variable
                        ParseEntities ForeignEntity, Run_Select()
                        Foreign = LoadedEntities(ForeignClass.Name)
                    end if

                    if Ubound(Foreign) > -1 then
                        if IsArray(PrimaryEntity(PrimaryField)) then
                            For Each Entity in Primary
                                ForeignIndex = -1
                                For Each ForeignEntity in Foreign
                                    if Entity(Field) = ForeignEntity(Field) then
                                        ForeignIndex = ForeignIndex + 1
                                    end if
                                Next
                                ForeignList = EmptyArray(ForeignIndex)
                                ForeignIndex = 0
                                For Each ForeignEntity in Foreign
                                    if Entity(Field) = ForeignEntity(Field) then
                                        set ForeignList(ForeignIndex) = ForeignEntity
                                        ForeignIndex = ForeignIndex + 1
                                    end if
                                Next
                                Entity(PrimaryField) = ForeignList
                            Next
                        else
                            For Each Entity in Primary
                                For Each ForeignEntity in Foreign
                                    if Entity(Field) = ForeignEntity(Field) then
                                        set Entity(PrimaryField) = ForeignEntity
                                        Exit For
                                    end if
                                Next
                            Next
                        end if
                    end if
                Next

                ' Recursion
                For Each PrimaryField in ForeignEntities
                    ForeignClass = ForeignEntities(PrimaryField)
                    if LoadedEntities.Exists(ForeignClass) then
                        if UBound(LoadedEntities(ForeignClass)) <> -1 then
                            Read_Foreign LoadedEntities(ForeignClass)
                        end if
                    end if
                Next
            end if
        End Sub
        ' Updates the $Entity's register  from it's table on Database
        '
        ' @param {object} Entity
        ' @return {int}
        Public Function Update( Entity )
            Dim Fields : Set Fields = EntityFields(Entity)
            Fields.Remove Entity.KeyField

            Call CurrentStatement _
                .Update_Clause(Entity.Self.Field("TableName")) _
                .Set_Clause(Fields)
            Where_Entity Entity

            Update = Run_Update()
        End Function
        ' Deletes the $Entity's register from it's table on Database
        '
        ' @param {object} Entity
        ' @return {int}
        Public Function Delete( Entity )
            CurrentStatement.From_Clause Entity.Self.Field("TableName")
            Where_Entity Entity

            Delete = Run_Delete()
        End Function
End Class

Dim DB_Instance : Set DB_Instance = new DataBase
%>