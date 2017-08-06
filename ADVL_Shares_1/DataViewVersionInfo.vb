Public Class DataViewVersionInfo
    'Class used to store different versions of the Data View.

    'LIST OF PROPERTIES:
    'VersionName            The version of the Data View.
    'VersionDesc        A description of the Version of the Data View.
    'Query              The query used to select the data from a database.
    'AutoApplyQuery     If True the Query is applied when the data view form is opened. 

    Public TableCols As New List(Of String) 'Stores a list of the table columns returned by the query.

    Private _versionName As String = "" 'The Name of the version of the Data View.
    Property VersionName As String
        Get
            Return _versionName
        End Get
        Set(value As String)
            _versionName = value
        End Set
    End Property

    Private _versionDesc As String = "" 'A description of the Version of the Data View.
    Property VersionDesc As String
        Get
            Return _versionDesc
        End Get
        Set(value As String)
            _versionDesc = value
        End Set
    End Property

    Private _query As String = "" 'The query used to select the data from a database.
    Property Query As String
        Get
            Return _query
        End Get
        Set(value As String)
            _query = value
        End Set
    End Property

    Private _autoApplyQuery As Boolean = False 'If True the Query is applied when the data view form is opened.
    Property AutoApplyQuery As Boolean
        Get
            Return _autoApplyQuery
        End Get
        Set(value As Boolean)
            _autoApplyQuery = value
        End Set
    End Property

End Class
