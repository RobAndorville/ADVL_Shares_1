Public Class DataViewSettingsList
    'Stores a list of data view settings.

    Public List As New List(Of DataViewSettings) 'List of Data View Settings.

    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation 'The location of the list file.

#Region " Properties" '===================================================================================================

    Private _listFileName As String = ""
    Property ListFileName As String 'The file name (with extension) of the list file.
        Get
            Return _listFileName
        End Get
        Set(value As String)
            _listFileName = value
        End Set
    End Property

    Private _creationDate As DateTime = Now 'The date of creation of the list.
    Property CreationDate As DateTime
        Get
            Return _creationDate
        End Get
        Set(value As DateTime)
            _creationDate = value
        End Set
    End Property

    Private _lastEditDate As DateTime = Now 'The last edit date of the list.
    Property LastEditDate As DateTime
        Get
            Return _lastEditDate
        End Get
        Set(value As DateTime)
            _lastEditDate = value
        End Set
    End Property

    Private _description As String = "" 'A description of the list.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _nRecords As Integer = 0 'The number of record in the list
    ReadOnly Property NRecords As Integer
        Get
            _nRecords = List.Count
            Return _nRecords
        End Get
    End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '=======================================================================================================

    'Clear the list.
    Public Sub Clear()
        List.Clear()
        ListFileName = ""
        Description = ""
    End Sub

    'Load the XML data in the XDoc into the Data View Settings List.
    Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

        CreationDate = XDoc.<SettingsList>.<CreationDate>.Value
        LastEditDate = XDoc.<SettingsList>.<LastEditDate>.Value
        Description = XDoc.<SettingsList>.<Description>.Value

        Dim Settings = From item In XDoc.<SettingsList>.<DataViewSettings>

        List.Clear()

        For Each item In Settings
            Dim NewSettings As New DataViewSettings
            NewSettings.Left = item.<Left>.Value
            NewSettings.Top = item.<Top>.Value
            NewSettings.Height = item.<Height>.Value
            NewSettings.Width = item.<Width>.Value
            If item.<VersionNo>.Value <> Nothing Then NewSettings.VersionNo = item.<VersionNo>.Value  'DataViewSettings did not originally store versions. This code allows older original files to be read.
            If item.<VersionName>.Value <> Nothing Then NewSettings.VersionName = item.<VersionName>.Value
            If item.<VersionDesc>.Value <> Nothing Then NewSettings.VersionDesc = item.<VersionDesc>.Value
            NewSettings.Query = item.<Query>.Value
            NewSettings.Description = item.<Description>.Value
            NewSettings.AutoApplyQuery = item.<AutoApplyQuery>.Value
            If item.<SelectedTab>.Value <> Nothing Then NewSettings.SelectedTab = item.<SelectedTab>.Value
            If item.<SaveFileDir>.Value <> Nothing Then NewSettings.SaveFileDir = item.<SaveFileDir>.Value
            If item.<XmlFileName>.Value <> Nothing Then NewSettings.XmlFileName = item.<XmlFileName>.Value
            For Each colItem In item.<ColumnList>.<Column>
                NewSettings.TableCols.Add(colItem)
            Next
            'If item.<VersionList>.<Version>.<Name>.Value = Nothing Then
            'If item.<VersionList>  <> Nothing Then
            For Each versItem In item.<VersionList>.<Version>
                    'NewSettings.Versions. = versItem.<Name>.Value
                    Dim NewVersion As New DataViewVersionInfo
                    NewVersion.VersionName = versItem.<Name>.Value
                    NewVersion.VersionDesc = versItem.<Description>.Value
                    NewVersion.Query = versItem.<Query>.Value
                    NewVersion.AutoApplyQuery = versItem.<AutoApplyQuery>.Value
                    For Each colItem In versItem.<ColumnList>
                        NewVersion.TableCols.Add(colItem)
                    Next
                    NewSettings.Versions.Add(NewVersion)
                Next
            'End If

            List.Add(NewSettings)
        Next

    End Sub

    'Load the list from the selected list file.
    Public Sub LoadFile()
        If ListFileName = "" Then 'No list file has been selected.
            RaiseEvent ErrorMessage("No list file name has been specified!" & vbCrLf)
            Exit Sub
        End If

        Dim XDoc As System.Xml.Linq.XDocument
        FileLocation.ReadXmlData(ListFileName, XDoc)

        If IsNothing(XDoc) Then
            RaiseEvent ErrorMessage("The specified list file contains no data!" & vbCrLf)
            Exit Sub
        End If

        LoadXml(XDoc)

    End Sub

    'Function to return the list of Data View Settings as an XDocument
    Public Function ToXDoc() As System.Xml.Linq.XDocument
        Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                   <!---->
                   <!--Data View Settings list file-->
                   <SettingsList>
                       <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                       <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                       <Description><%= Description %></Description>
                       <!---->
                       <%= From item In List
                           Select
                           <DataViewSettings>
                               <Left><%= item.Left %></Left>
                               <Top><%= item.Top %></Top>
                               <Height><%= item.Height %></Height>
                               <Width><%= item.Width %></Width>
                               <Description><%= item.Description %></Description>
                               <VersionNo><%= item.VersionNo %></VersionNo>
                               <VersionName><%= item.VersionName %></VersionName>
                               <VersionDesc><%= item.VersionDesc %></VersionDesc>
                               <Query><%= item.Query %></Query>
                               <AutoApplyQuery><%= item.AutoApplyQuery %></AutoApplyQuery>
                               <SelectedTab><%= item.SelectedTab %></SelectedTab>
                               <SaveFileDir><%= item.SaveFileDir %></SaveFileDir>
                               <XmlFileName><%= item.XmlFileName %></XmlFileName>
                               <ColumnList>
                                   <%= From colItem In item.TableCols
                                       Select
                                           <Column><%= colItem %></Column>
                                   %>
                               </ColumnList>
                               <VersionList>
                                   <%= From versItem In item.Versions
                                       Select
                                       <Version>
                                           <Name><%= versItem.VersionName %></Name>
                                           <Description><%= versItem.VersionDesc %></Description>
                                           <Query><%= versItem.Query %></Query>
                                           <AutoApplyQuery><%= versItem.AutoApplyQuery %></AutoApplyQuery>
                                           <ColumnList>
                                               <%= From colItem In versItem.TableCols
                                                   Select
                                                   <Column><%= colItem %></Column>
                                               %>
                                           </ColumnList>
                                       </Version> %>
                               </VersionList>
                           </DataViewSettings> %>
                   </SettingsList>
        Return XDoc
    End Function

    'Save the list in the selected list file.
    Public Sub SaveFile()
        If ListFileName = "" Then 'No list file has been selected.
            RaiseEvent ErrorMessage("No list file name has been specified!" & vbCrLf)
            Exit Sub
        End If

        FileLocation.SaveXmlData(ListFileName, ToXDoc)

    End Sub

    'Insert a Settings entry at the specified position in the list.
    Public Sub InsertSettings(ByVal Index As Integer, Item As DataViewSettings)
        'If Index + 1 = List.Count Then
        If Index = List.Count Then
            'Append the Settings to the end of the List:
            List.Add(Item)
            'ElseIf Index + 1 > List.Count Then
        ElseIf Index > List.Count Then
            RaiseEvent ErrorMessage("Index position is too large. Cannot insert the settings into the list." & vbCrLf)
        ElseIf Index < 0 Then
            RaiseEvent ErrorMessage("Index position is less than zero. Cannot insert the settings into the list." & vbCrLf)
        Else
            'Move existing entries to make space for the new settings:
            Dim LastIndex As Integer = List.Count - 1
            List.Add(List(LastIndex)) 'Append a copy of the last settings to the end of the list.
            Dim I As Integer
            For I = LastIndex To Index + 1 Step -1
                List(I) = List(I - 1)
            Next
            List(Index) = Item
        End If
    End Sub

    'Update a Settings entry at the specified position in the list.
    Public Sub UpdateSettings(ByVal Index As Integer, Item As DataViewSettings)
        If Index + 1 > List.Count Then
            RaiseEvent ErrorMessage("Index position is too large. Cannot modify the settings in the list." & vbCrLf)
        ElseIf Index < 0 Then
            RaiseEvent ErrorMessage("Index position is less than zero. Cannot modify the settings in the list." & vbCrLf)
        Else
            List(Index) = Item
        End If
    End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '========================================================================================================
    Event ErrorMessage(ByVal Message As String) 'Send an error message.
    Event Message(ByVal Message As String) 'Send a normal message.
#End Region 'Events ------------------------------------------------------------------------------------------------------

End Class
