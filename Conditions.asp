<%
' Basic condition structure.
' Mainly for WHERE clause.
Class DB_Condition
    ' Public interface
        ' Local operator to insert before this condition in a conditions list.
        '
        ' @var {string}
        Public Operator
        ' Field to check value(s) or simple fixed condition.
        ' Wich one of this two depends on the presence of
        ' $Values in this struct.
        '
        ' @var {string}
        Public Field
        ' Value(s) to check against $Field.
        ' For IN clauses will be an array.
        '
        ' @var {mixed}
        Public Values



    ' Real optional Constructor
        ' @param {string} Operator_
        ' @param {string} Field_
        ' @param {mixed} Values_
        ' @return {self}
        Public Default Function Construct( Operator_, Field_, Values_ )
            Operator = Operator_
            Field    = Field_
            Values   = Values_

            set Construct = Me
        End Function
End Class



' Multi-condition structure for JOIN clauses.
Class DB_JoinClause
    ' Public interface
        ' Basic DB_Condition container.
        '
        ' @var {Dictionary<int, Condition>}
        Public Conditions
        ' Join operator to insert before this join in a join-list.
        '
        ' @var {string}
        Public JoinType
        ' Table to join.
        '
        ' @var {string}
        Public Table



    ' Pre-constructor
        Sub Class_Initialize( )
            Set Conditions = CreateObject("Scripting.Dictionary")
        End Sub



    ' Real optional Constructor.
        ' @param {string} JoinType_
        ' @param {string} Table_
        ' @return {self}
        Public Default Function Construct( JoinType_, Table_ )
            JoinType = JoinType_
            Table    = Table_

            set Construct = Me
        End Function
    ' Destructor
        Sub Class_Terminate( )
            Conditions.RemoveAll()
            Set Conditions = Nothing
        End Sub
End Class
%>