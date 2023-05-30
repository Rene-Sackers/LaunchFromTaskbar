Module Core
    Public PathToFile As String = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Flamefusion\LaunchFromTaskbar\Configuration"
    Public FilePath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Flamefusion\LaunchFromTaskbar\Configuration\Items.txt"
    Public MainWin As MainWindow
    Public EditWin As New Editor

    Public ItemList As New List(Of String)

    Public Method As String = "Add"

    Private MyJumpList As New Shell.JumpList

    Public Sub CollectItems()
        Dim Reader As New IO.StreamReader(FilePath, System.Text.Encoding.Default)
        Dim Line As String

        Do While Reader.EndOfStream = False
            Try
                Line = Reader.ReadLine

                ItemList.Add(Line)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Loop

        Reader.Close() : Reader.Dispose()
    End Sub

    Public Sub Save()
        Dim Writer As New IO.StreamWriter(FilePath, False, System.Text.Encoding.Default)

        For Each S As String In ItemList
            Writer.WriteLine(S)
        Next

        Writer.Close() : Writer.Dispose()
    End Sub

    Public Sub SetJumpListItems()
        MyJumpList = New Shell.JumpList
        Shell.JumpList.SetJumpList(Application.Current, MyJumpList)
        MyJumpList.ShowRecentCategory = False
        MyJumpList.ShowFrequentCategory = False

        Dim Drives As String = Nothing
        For Each D As System.IO.DriveInfo In System.IO.DriveInfo.GetDrives
            Drives += " " & D.Name
        Next

        Dim Reader As New IO.StreamReader(FilePath, System.Text.Encoding.Default, False)
        Dim CurrentCatagory As String = Nothing
        Dim Line As String = Nothing
        Dim InNewCatagory As Boolean = False

        Do Until Reader.EndOfStream
            Line = Reader.ReadLine
            If Line.StartsWith(" catagory ") Then
                CurrentCatagory = Line.Remove(0, 10)
                InNewCatagory = True
            Else
                Dim Task As New Shell.JumpTask
                Task.CustomCategory = CurrentCatagory
                Task.Title = Line.Split(" "c)(0)
                Task.ApplicationPath = Line.Split(" "c)(1)
                If Not Line.Split(" "c).Count = 3 Then
                    If Not IO.Directory.Exists(Line.Split(" "c)(1)) And Not Line.Split(" "c)(1).StartsWith("http://") Then
                        If Line.Split(" "c)(1).EndsWith(".txt") Then
                            Task.IconResourcePath = "%windir%\System32\shell32.dll"
                            Task.IconResourceIndex = 70
                        ElseIf Line.Split(" "c)(1) = ("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") Then
                            Task.IconResourcePath = "%windir%\System32\shell32.dll"
                            Task.IconResourceIndex = 15
                        Else
                            Task.IconResourcePath = Line.Split(" "c)(1)
                        End If
                    ElseIf IO.Directory.Exists(Line.Split(" "c)(1)) Then
                        Try
                            If Drives.Contains(Line.Split(" "c)(1)) Then
                                Task.IconResourcePath = "%windir%\System32\shell32.dll"
                                Task.IconResourceIndex = 8
                            Else
                                Task.IconResourcePath = "%windir%\System32\shell32.dll"
                                Task.IconResourceIndex = 3
                            End If
                        Catch
                            Task.IconResourcePath = "%windir%\System32\shell32.dll"
                            Task.IconResourceIndex = 3
                        End Try
                    ElseIf Line.Split(" "c)(1).StartsWith("http://") Then
                            Task.IconResourcePath = "%windir%\System32\shell32.dll"
                            Task.IconResourceIndex = 135
                    End If
                Else
                    Task.IconResourcePath = "%windir%\System32\shell32.dll"
                    Task.IconResourceIndex = 0
                End If
                If InNewCatagory = True Then
                    MyJumpList.JumpItems.Insert(0, Task)
                    InNewCatagory = False
                Else
                    MyJumpList.JumpItems.Add(Task)
                End If
            End If
        Loop
        Reader.Close()
        Reader.Dispose()

        MyJumpList.Apply()
    End Sub

    Sub MySub(ByVal File As String)

        System.IO.File.Delete(File)

        System.IO.File.Create(File)

    End Sub

End Module
