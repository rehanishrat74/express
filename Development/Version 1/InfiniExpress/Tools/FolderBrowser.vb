Option Strict On

'Add a reference to "System.Design" (system.design.dll) to your project
Imports System.Windows.Forms.Design

Public Class FolderBrowser

    Inherits FolderNameEditor

    Public Enum enuFolderBrowserFolder
        Desktop = FolderBrowserFolder.Desktop
        Favorites = FolderBrowserFolder.Favorites
        MyComputer = FolderBrowserFolder.MyComputer
        MyDocuments = FolderBrowserFolder.MyDocuments
        MyPictures = FolderBrowserFolder.MyPictures
        NetAndDialUpConnections = FolderBrowserFolder.NetAndDialUpConnections
        NetworkNeighborhood = FolderBrowserFolder.NetworkNeighborhood
        Printers = FolderBrowserFolder.Printers
        Recent = FolderBrowserFolder.Recent
        SendTo = FolderBrowserFolder.SendTo
        StartMenu = FolderBrowserFolder.StartMenu
        Templates = FolderBrowserFolder.Templates
    End Enum
    Public StartLocation As enuFolderBrowserFolder = enuFolderBrowserFolder.MyComputer

    'The FolderBrowserStyles collection is a member of FolderNameEditor
    Public Enum enuFolderBrowserStyles
        BrowseForComputer = FolderBrowserStyles.BrowseForComputer
        BrowseForEverything = FolderBrowserStyles.BrowseForEverything
        BrowseForPrinter = FolderBrowserStyles.BrowseForPrinter
        RestrictToDomain = FolderBrowserStyles.RestrictToDomain
        RestrictToFilesystem = FolderBrowserStyles.RestrictToFilesystem
        RestrictToSubfolders = FolderBrowserStyles.RestrictToSubfolders
        ShowTextBox = FolderBrowserStyles.ShowTextBox
    End Enum
    Public Style As enuFolderBrowserStyles = enuFolderBrowserStyles.ShowTextBox

    Private mstrDescription As String = _
        "Please select a directory below:"
    Private mstrPath As String = String.Empty
    Private mobjFB As New FolderBrowser()

    Public Property Description() As String
        Get
            Return mstrDescription
        End Get
        Set(ByVal Value As String)
            mstrDescription = Value
        End Set
    End Property

    Public ReadOnly Property Path() As String
        Get
            Return mstrPath
        End Get
    End Property

    Public Function ShowBrowser() As System.Windows.Forms.DialogResult
        With mobjFB
            .Description = mstrDescription
            .StartLocation = CType(Me.StartLocation, FolderNameEditor.FolderBrowserFolder)
            .Style = CType(Me.Style, FolderNameEditor.FolderBrowserStyles)
            Dim dlgResult As DialogResult = .ShowDialog
            If dlgResult = DialogResult.OK Then
                mstrPath = .DirectoryPath
            Else
                mstrPath = String.Empty
            End If
            Return dlgResult
        End With
    End Function
End Class
