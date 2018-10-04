<%
' Crud Interface
    ' Gets the name of the field set to be used as primary key.
    ' If none is set, returns the first registered field.
    '
    ' @return {string}
    Public Property Get KeyField( )
        Dim Key : Key = Self.Field("KeyField")
        if IsEmpty(Key) then
            ' First field as key
            KeyField = Class_Loader.Members(Self).Keys()(0)
        else
            KeyField = Key
        end if
    End Property
    ' Inserts this object on it's database table.
    '
    ' @return {int}
    Public Function Create( )
        Create = DB_Instance.Create(Me)
    End Function
    ' Querys this object's database table.
    '
    ' @return {array<self>}
    Public Function Read( )
        Read = DB_Instance.Read(Me)
    End Function
    ' Updates this object on it's database table.
    '
    ' @return {int}
    Public Function Update( )
        Update = DB_Instance.Update(Me)
    End Function
    ' Deletes this object from it's database table.
    '
    ' @return {int}
    Public Function Delete( )
        Delete = DB_Instance.Delete(Me)
    End Function



' Queryable Interface
    ' If this Entity is set to be used for queries.
    ' @var {bool}
    Public Queryable



    ' Resets this object fields to their default values.
    '
    ' @return {self}
    Public Function ToDefaults()
        Call Instance_Initialize()

        Set ToDefaults = Me
    End Function
    ' Marks this object to not be used for queries.
    '
    ' @return {self}
    Public Function ToNonQueryable()
        Queryable = false

        Set ToNonQueryable = Me
    End Function
    ' Marks this object to be used in queries, setting all fields to empty.
    '
    ' @return {self}
    Public Function ToQueryable( )
        Queryable = true

        For Each Field_ in Class_Loader.Members(Self).Keys()
            Field(Field_) = Empty
        Next

        Set ToQueryable = Me
    End Function



' JSON export
    ' Exports this Entity to a JSONobject, adding all registered Foreign
    ' entities to it.
    '
    ' @return {JSONobject}
    Public Function ToJSON()
        Set ToJSON = Class_Loader.ToJSON(Me)
        if TypeName(Self.Field("Foreign")) = "Dictionary" then
            Dim Foreign : Set Foreign = Self.Field("Foreign")
            Dim Field_
            Dim Entity
            Dim List
            Dim Value

            For Each Field_ in Foreign
                set_ Value, Field(Field_) 
                if IsArray(Value) then
                    set List = new JSONarray
                    For Each Entity in Value
                        List.push Entity.ToJSON()
                    Next
                    ToJSON.Add Field_, List
                elseif not IsEmpty(Value) then
                    ToJSON.Add Field_, Value.ToJSON()
                end if
            Next
        end if
    End Function
    ' Exports this Entity to a JSON string, adding all registered Foreign
    ' entities to it.
    '
    ' @return {JSONobject}
    Public Function ToString()
        ToString = ToJSON().Serialize()
    End Function
%>