Imports System.IO
Imports System.Text

Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Windows.Forms


Public Class FileFolderDialog
    Inherits CommonDialog

    Private dialog As New OpenFileDialog

    'Public Overloads Property Dialog2 As OpenFileDialog
    '    Get
    '        Console.WriteLine("dialog test")
    '        Return Dialog2
    '    End Get
    '    Set(ByVal value As OpenFileDialog)
    '        Console.WriteLine("dialog test")
    '        Dialog2 = value
    '    End Set
    'End Property

    'Public Overloads Function Get() As DialogResult
    '    Console.WriteLine("dialog test3")

    'End Function


    Public Overloads Function ShowDialog() As DialogResult
        Console.WriteLine("dialog test3")
        Return Me.ShowDialog(Nothing)
    End Function

    Public Overloads Function ShowDialog(ByVal owner As IWin32Window) As DialogResult
        Console.WriteLine("dialog test4")
        Dialog.ValidateNames = False
        Dialog.CheckFileExists = False
        dialog.CheckPathExists = True

        Try

            If dialog.FileName IsNot Nothing AndAlso dialog.FileName <> "" Then

                If Directory.Exists(dialog.FileName) Then
                    dialog.InitialDirectory = dialog.FileName
                Else
                    dialog.InitialDirectory = Path.GetDirectoryName(dialog.FileName)
                End If
            End If

        Catch ex As Exception
        End Try

        dialog.FileName = "Folder Selection."

        If owner Is Nothing Then
            Return dialog.ShowDialog()
        Else
            Return dialog.ShowDialog(owner)
        End If
    End Function

    Public Property SelectedPath As String
        Get

            Try

                If dialog.FileName IsNot Nothing AndAlso (dialog.FileName.EndsWith("Folder Selection.") OrElse Not File.Exists(dialog.FileName)) AndAlso Not Directory.Exists(dialog.FileName) Then
                    Return Path.GetDirectoryName(dialog.FileName)
                Else
                    Return dialog.FileName
                End If

            Catch ex As Exception
                Return dialog.FileName
            End Try
        End Get
        Set(ByVal value As String)

            If value IsNot Nothing AndAlso value <> "" Then
                dialog.FileName = value
            End If
        End Set
    End Property

    Public ReadOnly Property SelectedPaths As String
        Get

            If dialog.FileNames IsNot Nothing AndAlso dialog.FileNames.Length > 1 Then
                Dim sb As StringBuilder = New StringBuilder()

                For Each fileName As String In dialog.FileNames

                    Try
                        If File.Exists(fileName) Then sb.Append(fileName & ";")
                    Catch ex As Exception
                    End Try
                Next

                Return sb.ToString()
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overrides Sub Reset()
        dialog.Reset()
    End Sub

    Protected Overrides Function RunDialog(ByVal hwndOwner As IntPtr) As Boolean
        Console.WriteLine("dialog test5")
        Return True
    End Function


End Class
