Public Class TableColumn
    Private _ColumnName As String
    Private _TableName As String
    Private _DataType As DataType
    Private _data_type As String
    Private _MaxCharacterLength As Integer
    Private _IsNullable As Boolean

    Public Enum DataType
        BigInt
        Int
        Varchar
        TinyInt
        SmallInt
        [Decimal]
        DateTime
        Bit
        [Char]
        Float
        BinaryArray
    End Enum

    Public ReadOnly Property ColumnName() As String
        Get
            Return _ColumnName
        End Get
    End Property

    Public ReadOnly Property ColumnNameFormatted() As String
        Get
            If IsVBReservedWord Then
                Return "[" & _ColumnName & "]"
            Else
                Return _ColumnName
            End If

        End Get
    End Property

    Public ReadOnly Property TableName() As String
        Get
            Return _TableName
        End Get
    End Property

    Public ReadOnly Property [Type]() As DataType
        Get
            Return _DataType
        End Get
    End Property

    Public ReadOnly Property MaxCharacterLength() As Integer
        Get
            Return _MaxCharacterLength
        End Get
    End Property

    Public ReadOnly Property IsNullable() As Boolean
        Get
            Return _IsNullable
        End Get
    End Property

    Public Sub New(ByVal columnName As String, ByVal tableName As String, _
                   ByVal [type] As String, ByVal maxCharLength As Integer, ByVal isNullable As String)
        _ColumnName = columnName
        _TableName = tableName
        Select Case [type]
            Case "varchar", "nvarchar" : _DataType = DataType.Varchar
            Case "bigint" : _DataType = DataType.BigInt
            Case "int" : _DataType = DataType.Int
            Case "smallint" : _DataType = DataType.SmallInt
            Case "tinyint" : _DataType = DataType.TinyInt
            Case "decimal" : _DataType = DataType.Decimal
            Case "datetime", "smalldatetime" : _DataType = DataType.DateTime
            Case "char" : _DataType = DataType.Char
            Case "bit" : _DataType = DataType.Bit
            Case "float" : _DataType = DataType.Float
            Case "varbinary" : _DataType = DataType.BinaryArray
            Case Else : Throw New ApplicationException("Unknown type: " & [type])
        End Select
        _MaxCharacterLength = maxCharLength

        If isNullable.ToUpper.Trim = "YES" Then
            _IsNullable = True
        Else
            _IsNullable = False
        End If
    End Sub

    Public Function VBColumnType() As String
        Dim res = ""

        Select Case _DataType
            Case DataType.Varchar : Return "String"
            Case DataType.BigInt : res = "Int64"
            Case DataType.Bit : res = "Boolean"
            Case DataType.Char : Return "String"
            Case DataType.DateTime : res = "DateTime"
            Case DataType.Int : res = "Integer"
            Case DataType.SmallInt : res = "Int16"
            Case DataType.TinyInt : res = "Byte"
            Case DataType.Decimal : res = "Decimal"
            Case DataType.Float : res = "Double"
            Case DataType.BinaryArray : res = "Byte()"
            Case Else : Throw New ApplicationException("Unknown type: " & [Type])
        End Select

        If _IsNullable Then Return res & "?"
        Return res
    End Function

    Public Function IsVBReservedWord() As Boolean
        Select Case _ColumnName.ToLower
            Case "addhandler", "addressof", "alias", "and", "andalso", "as", "boolean", "byref", "byte", _
                 "byval", "call", "case", "catch", "cbool", "cbyte", "cchar", "cdate", "cdec", "cdbl", "char", _
                 "cint", "class", "clng", "cobj", "const", "continue", "csbyte", "cshort", "csng", "cstr", _
                 "ctype", "cuint", "culng", "cushort", "date", "decimal", "declare", "default", "delegate", _
                 "dim", "directcast", "do", "double", "each", "else", "elseif", "end", "endif", "enum", _
                 "erase", "error", "event", "exit", "false", "finally", "for", "friend", "function", "get", _
                 "gettype", "getxmlnamespace", "global", "gosub", "goto", "handles", "if", "if()", _
                 "implements", "imports (.net namespace and type)", "imports (xml namespace)", "in", _
                 "inherits", "integer", "interface", "is", "isnot", "let", "lib", "like", "long", "loop", _
                 "me", "mod", "module", "mustinherit", "mustoverride", "mybase", "myclass", "namespace", _
                 "narrowing", "new", "next", "not", "nothing", "notinheritable", "notoverridable", "object", _
                 "of", "on", "operator", "option", "optional", "or", "orelse", "overloads", "overridable", _
                 "overrides", "paramarray", "partial", "private", "property", "protected", "public", _
                 "raiseevent", "readonly", "redim", "rem", "removehandler", "resume", "return", "sbyte", _
                 "select", "set", "shadows", "shared", "short", "single", "static", "step", "stop", "string", _
                 "structure", "sub", "synclock", "then", "throw", "to", "true", "try", "trycast", "typeof", _
                 "variant", "wend", "uinteger", "ulong", "ushort", "using", "when", "while", "widening", _
                 "with", "withevents", "writeonly", "xor"
                Return True
            Case Else
                Return False
        End Select
    End Function
End Class
