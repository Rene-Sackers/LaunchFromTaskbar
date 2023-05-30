Class MainWindow

    Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        MainWin = Me
        CollectItems()
        SetJumpListItems()
    End Sub

    Private Sub AddButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles AddButton.Click
        Method = "Add"
        EditWin.Window_Loaded()
        Me.Hide()
        EditWin.Show()
    End Sub

    Private Sub EditButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles EditButton.Click
        Method = "Edit"
        EditWin.Window_Loaded()
        Me.Hide()
        EditWin.Show()
    End Sub

    Private Sub MoveButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MoveButton.Click
        Method = "Move"
        EditWin.Window_Loaded()
        Me.Hide()
        EditWin.Show()
    End Sub

    Private Sub RemoveButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles RemoveButton.Click
        Method = "Remove"
        EditWin.Window_Loaded()
        Me.Hide()
        EditWin.Show()
    End Sub

    Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        End
    End Sub
End Class
