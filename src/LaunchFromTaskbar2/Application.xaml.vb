Class Application

    Private Sub Application_Startup(ByVal sender As Object, ByVal e As System.Windows.StartupEventArgs) Handles Me.Startup

        If IO.Directory.Exists(PathToFile) = False Then

            IO.Directory.CreateDirectory(PathToFile)

        End If

        If IO.File.Exists(FilePath) = False Then

            Dim StreamWriter As IO.StreamWriter = IO.File.CreateText(FilePath)
            StreamWriter.Close()

        End If

    End Sub
End Class
