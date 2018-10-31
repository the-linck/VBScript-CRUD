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
        Instance_Initialize()

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



' Import
    ' Creates/feeds Entities with data present on given Source.
    '
    ' @param {Scripting.Dictionary} Source
    ' @return {Object}
    Public Function FromDictionary(Source)
        Class_Loader.FromDictionary Me, Source

        if TypeName(Self.Field("Foreign")) = "Dictionary" then
            Dim Foreign : Set Foreign = Self.Field("Foreign")
            Dim Field_
            Dim Entity
            Dim Index
            Dim Key
            Dim List
            Dim Value

            For Each Field_ in Foreign
                set_ Value, Source(Field_)
                if TypeName(Value) = "Dictionary" then
                    ' Getting only first element to check whole list type
                    For Each Key in Value
                        Exit For
                    Next

                    if IsNumeric(Key) then ' Arrays use numeric keys
                        List = EmptyArray(Value.Count - 1)
                        Index = -1
                        For Each Key in Value
                            set Entity = Class_Loader.FromDictionary(Foreign(Field), Value(Key))
                            if not Entity Is Nothing then
                                Set List(Index) = Entity
                                Index = Index + 1
                            end if
                        Next
                        Redim Preserve List(Index)
                        Field(Field_) = List
                    else ' Entities use string keys
                        set Entity = Class_Loader.FromDictionary(Foreign(Field), Value)
                        if not Entity Is Nothing then
                            Field(Field_) = Entity
                        end if
                    end if
                end if
            Next
        end if

        Set FromDictionary = Me
    End Function
    ' Creates/feeds Entities with data present on given Source.
    '
    ' @param {JSONobject|JSONarray|string} Source
    ' @return {Object|Object[]}
    Public Function FromJSON(Source)
        Dim JSON
        if TypeName(Source) = "JSONobject" or TypeName(Source) = "JSONarray" then
            set JSON = Source
        else
            set JSON = (new JSONobject).parse(Source)
        end if
        set_ FromJSON, Class_Loader.FromJSON(Me, JSON)

        if TypeName(Self.Field("Foreign")) = "Dictionary" then
            Dim Entity
            Dim EntityClass
            Dim Foreign
            Dim Field_
            Dim Index
            Set Foreign = Self.Field("Foreign")

            Select Case TypeName(JSON)
                ' Avoiding unecessary increase of the call stack
                Case "JSONarray"
                    ' Avoiding new object (and inloop verification)
                    For Each Field_ in Foreign
                        Set EntityClass = Class_Loader.Load(Foreign(Field_))
                        set Entity = EntityClass.GetInstance().FromJSON(JSON(0)(Field_))
                        if not Entity Is Nothing then
                            FromJSON(0).Field(Field_) = Entity
                        end if
                    Next

                    For Index = JSON.length - 1 To 1 Step -1
                        For Each Field_ in Foreign
                            Set EntityClass = Class_Loader.Load(Foreign(Field_))
                            set Entity = EntityClass.GetInstance().FromJSON(JSON(Index)(Field_))
                            if not Entity Is Nothing then
                                FromJSON(Index).Field(Field_) = Entity
                            end if
                        Next
                    Next
                Case "JSONobject"
                    For Each Field_ in Foreign
                        Set EntityClass = Class_Loader.Load(Foreign(Field_))
                        set Entity = EntityClass.GetInstance().FromJSON(JSON(Field_))
                        if not Entity Is Nothing then
                            Field(Field_) = Entity
                        end if
                    Next
            End Select
        end if
    End Function
    ' Creates/feeds Entities with data present on given request Method.
    ' Uses giver Prefix to identify fields names.
    '
    ' @param {string} Method [Form|Post|Querystring|Get]
    ' @return {Object}
    Public Function FromRequest(Method)
        Class_Loader.FromRequest Me, Method, ""

        if TypeName(Self.Field("Foreign")) = "Dictionary" then
            Dim Foreign : Set Foreign = Self.Field("Foreign")
            Dim Field_
            Dim Entity

            For Each Field_ in Foreign
                set Entity = Class_Loader.FromRequest(Foreign(Field), Method, Field & ".")

                if not Entity Is Nothing then
                    Field(Field_) = Entity
                end if
            Next
        end if

        Set FromRequest = Me
    End Function
    ' Creates/feeds Entities with a JSON string present on session Key.
    '
    ' @param {string} Key
    ' @return {Object}
    Public Function FromSession(Key)
        set_ FromSession, FromJSON(Session(Key))
    End Function
    ' Creates/feeds Entities with data present on given Source.
    '
    ' @param {string} Source
    ' @return {Object}
    Public Function FromString(Source)
        set_ FromString, FromJSON(Source)
    End Function



' Export
    ' Exports this Entity to a Dictionary.
    '
    ' @return {Scripting.Dictionary}
    Public Property Get ToDictionary()
        Set ToDictionary = Class_Loader.ToDictionary(Me)
        if TypeName(Self.Field("Foreign")) = "Dictionary" then
            Dim Foreign : Set Foreign = Self.Field("Foreign")
            Dim Field_
            Dim Entity
            Dim List
            Dim Value

            For Each Field_ in Foreign
                set_ Value, Field(Field_)
                if IsArray(Value) then
                    set List = CreateObject("Scripting.Dictionary")
                    For Each Entity in Value
                        List(List.Count) = Entity.ToDictionary()
                    Next
                    ToDictionary.Add Field_, List
                elseif IsObject(Value) then
                    if property_exists(Value, "SupportsReflection") then
                        set List = Value.ToDictionary()
                        ToDictionary.Add Field_, List
                    end if
                end if
            Next
        end if
    End Property
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
                elseif TypeName(Value) = "JSONobject" or TypeName(Value) = "JSONarray" then
                    ToJSON.Add Field_, Value
                elseif IsObject(Value) then
                    if property_exists(Value, "SupportsReflection") then
                        ToJSON.Add Field_, Value.ToJSON()
                    end if
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