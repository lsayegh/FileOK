Imports System.IO
Imports System.IO.Directory
Imports System.Data
Imports System.Data.SqlClient

Public Class frmMain


    '0) ParentDirectoryName
    '1) DirectoryName
    '2) DirectorySize
    '3) DirectoryDateCreatedUtc
    '4) DirectoryDateLastWriteUtc
    '5) DirectoryDateLastAccessUtc
    '6) FileName
    '7) FileExtension
    '8) FileSize
    '9) FileDateCreated
    '10) FileLastTimeAccessUtc
    '11) FileLastTimeWriteUtc
    '12) FileHashCode
    '13) IsReadOnly
    '14) Message

    Dim files(,) As String
    Dim dParentFullName As String
    Dim dFullName As String
    Dim dCreationTimeUtc As String
    Dim dLastAccessTimeUtc As String
    Dim dLastWriteTimeUtc As String


    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Dim locations() As String

        ReDim locations(3)
        locations(1) = "C:\"
        locations(2) = "F:\"
        locations(3) = "E:\"

        'Loop through each location element
        For Each item As String In locations
            Debug.Write(item & " ")
            Label1.Text = item
            If Not IsNothing(item) Then
                SearchFiles(item)
            End If
        Next
        Erase files

    End Sub

    Public Sub SearchFiles(ByVal SourcePath As String)
        Dim SourceDir As DirectoryInfo = New DirectoryInfo(SourcePath)
        Dim fFileName As String
        Dim fFileExtension As String
        Dim fFileSize As String
        Dim fFileDateCreated As String
        Dim fFileLastTimeAccessUtc As Date
        Dim fFileLastTimeWriteUtc As Date
        Dim fFileHashCode As String
        Dim fIsReadOnly As Boolean
        Dim message As String

        Try
            If SourceDir.Exists Then
                Dim SubDir As DirectoryInfo
                For Each SubDir In SourceDir.GetDirectories()
                    dFullName = SubDir.FullName
                    dParentFullName = SubDir.Parent.Name
                    dCreationTimeUtc = SubDir.CreationTimeUtc
                    dLastAccessTimeUtc = SubDir.LastAccessTimeUtc
                    dLastWriteTimeUtc = SubDir.LastWriteTimeUtc
                    Console.WriteLine("Folder Name: " & dFullName)
                    If Directory.GetFiles(dFullName).Count > 0 Then
                        SearchFiles(SubDir.FullName)
                    Else
                        AddFileDatabase(dParentFullName, dFullName, vbNull, dCreationTimeUtc,
                                dLastWriteTimeUtc, dLastAccessTimeUtc, , , , , , , , )
                    End If

                Next
                For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.AllDirectories)
                    fFileName = childFile.FullName
                    fFileExtension = childFile.Extension
                    fFileSize = childFile.Length
                    fFileDateCreated = childFile.CreationTimeUtc
                    fFileLastTimeAccessUtc = childFile.LastAccessTimeUtc
                    fFileLastTimeWriteUtc = childFile.LastWriteTimeUtc
                    fFileHashCode = childFile.GetHashCode
                    fIsReadOnly = childFile.IsReadOnly
                    Console.WriteLine("File Name: " & fFileName)
                    AddFileDatabase(dParentFullName, dFullName, vbNull, dCreationTimeUtc,
                        dLastWriteTimeUtc, dLastAccessTimeUtc, fFileName, fFileExtension, fFileSize, fFileDateCreated,
                        fFileLastTimeAccessUtc, fFileLastTimeWriteUtc, fFileHashCode, fIsReadOnly)

                Next
            End If

        Catch ex As UnauthorizedAccessException
            message = ex.Message + ex.StackTrace
            'AddFileArray(dParentFullName, dFullName, vbNull, dCreationTimeUtc,
            '            dLastWriteTimeUtc, dLastAccessTimeUtc, fFileName, fFileExtension, fFileSize, fFileDateCreated,
            '            fFileLastTimeAccessUtc, fFileLastTimeWriteUtc, fFileHashCode, fIsReadOnly, message)

        Catch ex As Exception
            message = ex.Message + ex.StackTrace
            'AddFileArray(dParentFullName, dFullName, vbNull, dCreationTimeUtc,
            '            dLastWriteTimeUtc, dLastAccessTimeUtc, fFileName, fFileExtension, fFileSize, fFileDateCreated,
            '            fFileLastTimeAccessUtc, fFileLastTimeWriteUtc, fFileHashCode, fIsReadOnly, message)
        End Try

    End Sub

    Private Sub AddFileArray(ByVal vParentDirectoryName As String, ByVal vDirectoryName As String, ByVal vDirectorySize As String, ByVal vDirectoryDateCreatedUtc As String,
                        ByVal vDirectoryDateLastWriteUtc As String, ByVal vDirectoryDateLastAccessUtc As String, ByVal vFileName As String,
                        ByVal vFileExtension As String, ByVal vFileSize As String, ByVal vFileDateCreated As String,
                        ByVal vFileLastTimeAccessUtc As String, ByVal vFileLastTimeWriteUtc As String,
                        ByVal vFileHashCode As String, ByVal vIsReadOnly As String, ByVal vMessage As String)
        Try
            Dim asize As Integer

            If files Is Nothing Then
                asize = 1
            Else
                asize = files.Rank + 1
            End If
            ReDim Preserve files(14, asize)

            files(1, asize) = vParentDirectoryName
            files(2, asize) = vDirectoryName
            files(3, asize) = vDirectorySize
            files(4, asize) = vDirectoryDateCreatedUtc
            files(5, asize) = vDirectoryDateLastWriteUtc
            files(6, asize) = vDirectoryDateLastAccessUtc
            files(7, asize) = IIf(vFileName Is Nothing, "Null", vFileName)
            files(8, asize) = vFileExtension
            files(9, asize) = vFileSize
            files(10, asize) = vFileDateCreated
            files(11, asize) = vFileLastTimeAccessUtc
            files(12, asize) = vFileLastTimeWriteUtc
            files(13, asize) = vFileHashCode
            files(14, asize) = vIsReadOnly
            files(15, asize) = vMessage

        Catch ex As Exception
            MsgBox(ex.Message + ex.StackTrace)
        End Try


    End Sub

    Private Sub AddFileDatabase(ByVal vParentDirectoryName As String, ByVal vDirectoryName As String, ByVal vDirectorySize As String, ByVal vDirectoryDateCreatedUtc As Date,
                        ByVal vDirectoryDateLastWriteUtc As Date, ByVal vDirectoryDateLastAccessUtc As Date, Optional ByVal vFileName As String = Nothing,
                        Optional ByVal vFileExtension As String = Nothing, Optional ByVal vFileSize As String = Nothing, Optional ByVal vFileDateCreated As Date = Nothing,
                        Optional ByVal vFileLastTimeAccessUtc As Date = Nothing, Optional ByVal vFileLastTimeWriteUtc As Date = Nothing,
                        Optional ByVal vFileHashCode As String = Nothing, Optional ByVal vIsReadOnly As Boolean = False)
        Dim cmd As New SqlCommand
        Dim conn As SqlConnection
        Dim daStudent As New SqlDataAdapter
        Dim Sql As String

        Try
            Sql = "INSERT INTO dbo.FilesInventory (DirectoryName,ParentDirectoryName,DirectoryDateCreatedUtc,DirectoryDateLastAccessUtc,DirectoryDateLastWriteUtc,FileNamefull,FileExtension,FileSize,FileDateCreatedUtc,FileDateLastAccessUtc,FileDateLastWriteUtc,FileHashCode,FileReadOnly)
     VALUES (@DirectoryName,@ParentDirectoryName,@DirectoryDateCreatedUtc,@DirectoryDateLastAccessUtc,@DirectoryDateLastWriteUtc,@FileNamefull,@FileExtension,@FileSize,@FileDateCreatedUtc,@FileDateLastAccessUtc,@FileDateLastWriteUtc,@FileHashCode,@FileReadOnly)"

            conn = New SqlConnection("server = .\SQL2016DEV;database = FileOK;Trusted_Connection = yes")
            conn.Open()
            cmd = conn.CreateCommand
            With cmd
                .CommandText = Sql
                .Parameters.AddWithValue("@DirectoryName", vDirectoryName)
                .Parameters.AddWithValue("@ParentDirectoryName", vParentDirectoryName)
                .Parameters.AddWithValue("@DirectoryDateCreatedUtc", IIf(IsDate(vDirectoryDateCreatedUtc), vDirectoryDateCreatedUtc, DBNull.Value))
                .Parameters.AddWithValue("@DirectoryDateLastAccessUtc", IIf(IsDate(vDirectoryDateLastAccessUtc), vDirectoryDateLastAccessUtc, DBNull.Value))
                .Parameters.AddWithValue("@DirectoryDateLastWriteUtc", IIf(IsDate(vDirectoryDateLastWriteUtc), vDirectoryDateLastWriteUtc, DBNull.Value))
                .Parameters.AddWithValue("@FileNamefull", vFileName)
                .Parameters.AddWithValue("@FileExtension", vFileExtension)
                .Parameters.AddWithValue("@FileSize", vFileSize)
                .Parameters.AddWithValue("@FileDateCreatedUtc", IIf(vFileDateCreated <> "#1/1/0001 12:00:00 AM#", vFileDateCreated, DBNull.Value))
                .Parameters.AddWithValue("@FileDateLastAccessUtc", IIf(vFileLastTimeAccessUtc <> "#1/1/0001 12:00:00 AM#", vFileLastTimeAccessUtc, DBNull.Value))
                .Parameters.AddWithValue("@FileDateLastWriteUtc", IIf(vFileLastTimeWriteUtc <> "#1/1/0001 12:00:00 AM#", vFileLastTimeWriteUtc, DBNull.Value))
                .Parameters.AddWithValue("@FileHashCode", vFileHashCode)
                .Parameters.AddWithValue("@FileReadOnly", vIsReadOnly)
                .ExecuteNonQuery()
            End With
            conn.Close()

        Catch ex As Exception
            conn.Close()
        End Try

    End Sub


End Class
