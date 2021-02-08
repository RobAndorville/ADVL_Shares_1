Public Class DataViewSettings
    'Class used to store the Settings of a Data View.

    Public TableCols As New List(Of String) 'Stores a list of the table columns returned by the query.

    Public Versions As New List(Of DataViewVersionInfo) 'Stores a list of other versions of the Data View.

    'LIST OF PROPERTIES:
    'Left               Stores the position of the Left of a form.
    'Top                Stores the position of the Top of a form.
    'Width              Stores the Width of a form.
    'Height             Stores the Height of the form
    'Description        A description of a data view.
    'Version            The version of the Data View.
    'VersionDesc        A description of the Version of the Data View.
    'Query              The query used to select the data from a database.
    'AutoApplyQuery     If True the Query is applied when the data view form is opened. 

    'The database path is not stored as the same data view parameters may be used on different version of the databases.
    'FormNo             The instance number of the form. THIS IS NOW INFERRED FROM THE SETTINGS LIST.

    Private _left As Integer = 10 'Stores the position of the Left of a form.
    Property Left As Integer
        Get
            Return _left
        End Get
        Set(value As Integer)
            _left = value
        End Set
    End Property

    Private _top As Integer = 10 'Stores the position of the Top of a form.
    Property Top As Integer
        Get
            Return _top
        End Get
        Set(value As Integer)
            _top = value
        End Set
    End Property

    Private _width As Integer = 1000 'Stores the Width of a form.
    Property Width As Integer
        Get
            Return _width
        End Get
        Set(value As Integer)
            _width = value
        End Set
    End Property

    Private _height As Integer = 500 'Stores the Height of the form
    Property Height As Integer
        Get
            Return _height
        End Get
        Set(value As Integer)
            _height = value
        End Set
    End Property

    'NOTE: FORM NUMBER IS INFERRED FROM THE SETTINGS LIST.
    'Private _formNo As Integer = -1 'The instance number of the form.
    'Property FormNo As Integer
    '    Get
    '        Return _formNo
    '    End Get
    '    Set(value As Integer)
    '        _formNo = value
    '    End Set
    'End Property

    Private _description As String = "" 'A description of a data view.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _versionNo As Integer = 0 'The number of the selected version of the Data View.
    Property VersionNo As Integer
        Get
            Return _versionNo
        End Get
        Set(value As Integer)
            _versionNo = value
        End Set
    End Property

    Private _versionName As String = "" 'The Name of the selected version of the Data View.
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

    Private _selectedTab As Integer = 0 'The selected Query or Information tab.
    Property SelectedTab As Integer
        Get
            Return _selectedTab
        End Get
        Set(value As Integer)
            _selectedTab = value
        End Set
    End Property

    Private _saveFileDir As String = "" 'The directory used to save XML files.
    Property SaveFileDir As String
        Get
            Return _saveFileDir
        End Get
        Set(value As String)
            _saveFileDir = value
        End Set
    End Property

    Private _xmlFileName As String = "" 'The name of an XML file to contain saved data.
    Property XmlFileName As String
        Get
            Return _xmlFileName
        End Get
        Set(value As String)
            _xmlFileName = value
        End Set
    End Property

    Private _chartSettingsFile As String = "" 'The name of the chart settings file.
    Property ChartSettingsFile As String
        Get
            Return _chartSettingsFile
        End Get
        Set(value As String)
            _chartSettingsFile = value
        End Set
    End Property

End Class
