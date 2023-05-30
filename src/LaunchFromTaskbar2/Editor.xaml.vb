Public Class Editor
    Private Browser As New Forms.OpenFileDialog

    Public Sub Window_Loaded() Handles MyBase.Loaded
        Me.Title = Method
        If Method = "Add" Then
            Me.Height = 180
            IndicatorLabel.Content = "Add Under:"
            NameTextBox.Text = ""
            PathTextBox.Text = ""
            NameTextBox.IsReadOnly = False
            PathTextBox.IsReadOnly = False
            BrowseButton.IsEnabled = True
            BrowseButton.Visibility = Visibility.Visible
            NameLabel.Visibility = Visibility.Visible
            PathLabel.Visibility = Visibility.Visible
            NameTextBox.Visibility = Visibility.Visible
            PathTextBox.Visibility = Visibility.Visible
            AddInfoLabel.Visibility = Visibility.Visible
            MoveLabel.Visibility = Visibility.Hidden
            MoveComboBox.Visibility = Visibility.Hidden
            UseUnknownFileIconCheckBox.Visibility = Visibility.Visible
        ElseIf Method = "Edit" Then
            Me.Height = 180
            IndicatorLabel.Content = "Edit:"
            NameTextBox.IsReadOnly = False
            PathTextBox.IsReadOnly = False
            BrowseButton.IsEnabled = True
            BrowseButton.Visibility = Visibility.Visible
            NameLabel.Visibility = Visibility.Visible
            PathLabel.Visibility = Visibility.Visible
            NameTextBox.Visibility = Visibility.Visible
            PathTextBox.Visibility = Visibility.Visible
            AddInfoLabel.Visibility = Visibility.Visible
            MoveLabel.Visibility = Visibility.Hidden
            MoveComboBox.Visibility = Visibility.Hidden
            UseUnknownFileIconCheckBox.Visibility = Visibility.Visible
        ElseIf Method = "Remove" Then
            Me.Height = 180
            IndicatorLabel.Content = "Remove:"
            NameTextBox.IsReadOnly = True
            PathTextBox.IsReadOnly = True
            BrowseButton.IsEnabled = False
            BrowseButton.Visibility = Visibility.Visible
            NameLabel.Visibility = Visibility.Visible
            PathLabel.Visibility = Visibility.Visible
            NameTextBox.Visibility = Visibility.Visible
            PathTextBox.Visibility = Visibility.Visible
            AddInfoLabel.Visibility = Visibility.Hidden
            MoveLabel.Visibility = Visibility.Hidden
            MoveComboBox.Visibility = Visibility.Hidden
            UseUnknownFileIconCheckBox.Visibility = Visibility.Hidden
        ElseIf Method = "Move" Then
            Me.Height = 146
            IndicatorLabel.Content = "Move:"
            NameTextBox.IsReadOnly = True
            PathTextBox.IsReadOnly = True
            BrowseButton.Visibility = Visibility.Hidden
            NameLabel.Visibility = Visibility.Hidden
            PathLabel.Visibility = Visibility.Hidden
            NameTextBox.Visibility = Visibility.Hidden
            PathTextBox.Visibility = Visibility.Hidden
            MoveLabel.Visibility = Visibility.Visible
            MoveComboBox.Visibility = Visibility.Visible
            AddInfoLabel.Visibility = Visibility.Hidden
            UseUnknownFileIconCheckBox.Visibility = Visibility.Hidden
        End If

        ItemsComboBox.Items.Clear()
        MoveComboBox.Items.Clear()
        If Not (Method = "Edit" Or Method = "Remove" Or Method = "Move") Then ItemsComboBox.Items.Add("")
        If Method = "Move" Then MoveComboBox.Items.Add("")
        For Each S As String In ItemList
            If S.StartsWith(" catagory ") Then
                ItemsComboBox.Items.Add("--" & S.Remove(0, 10) & "--")
                If Method = "Move" Then MoveComboBox.Items.Add("--" & S.Remove(0, 10) & "--")
            Else
                ItemsComboBox.Items.Add(S.Split(" "c)(0))
                If Method = "Move" Then MoveComboBox.Items.Add(S.Split(" "c)(0))
            End If
            If Method = "Move" Then MoveComboBox.SelectedIndex = MoveComboBox.Items.Count - 1
        Next
        ItemsComboBox.SelectedIndex = ItemsComboBox.Items.Count - 1

        CloseAfterOKCheckbox.IsChecked = CBool(GetSetting("LaunchFromTaskbar2", "Settings", "CloseAfterOKCheckbox", True))
    End Sub

    Private Sub ItemsComboBox_SelectionChanged() Handles ItemsComboBox.SelectionChanged
        If Not Method = "Move" And Not Method = "Add" Then
            Try
                If ItemList.Item(ItemsComboBox.SelectedIndex).StartsWith(" catagory ") Then
                    UseUnknownFileIconCheckBox.Visibility = Visibility.Hidden
                    NameTextBox.Text = ItemList.Item(ItemsComboBox.SelectedIndex).Remove(0, 10)
                    PathTextBox.Text = ""
                Else
                    If (Method = "Edit" Or Method = "Add") Then UseUnknownFileIconCheckBox.Visibility = Visibility.Visible
                    If ItemList.Item(ItemsComboBox.SelectedIndex).Split(" "c).Count = 3 And (Method = "Edit" Or Method = "Add") Then
                        UseUnknownFileIconCheckBox.IsChecked = True
                    Else
                        UseUnknownFileIconCheckBox.IsChecked = False
                    End If
                    NameTextBox.Text = ItemsComboBox.SelectedItem
                    PathTextBox.Text = ItemList.Item(ItemsComboBox.SelectedIndex).Split(" "c)(1)
                End If
            Catch
            End Try
        End If
    End Sub

    Private Sub CancelButton_Click() Handles CancelButton.Click
        Me.Hide()
        MainWin.Show()
    End Sub

    Private Sub Editor_Closing() Handles Me.Closing
        End
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Try
            If Method = "Add" Then

                If Not PathTextBox.Text.StartsWith("http://") Then
                    If Not PathTextBox.Text = "" And (System.IO.File.Exists(PathTextBox.Text) Or System.IO.Directory.Exists(PathTextBox.Text) Or PathTextBox.Text.ToLower = "computer" Or PathTextBox.Text.ToLower = "my computer") Then
                        If UseUnknownFileIconCheckBox.IsChecked = True Then
                            If PathTextBox.Text.ToLower = "computer" Or PathTextBox.Text.ToLower = "my computer" Then
                                ItemList.Insert(ItemsComboBox.SelectedIndex, NameTextBox.Text & " ::{20D04FE0-3AEA-1069-A2D8-08002B30309D} UseUnknownFileIcon")
                            Else
                                ItemList.Insert(ItemsComboBox.SelectedIndex, NameTextBox.Text & " " & PathTextBox.Text & " UseUnknownFileIcon")
                            End If
                        Else
                            If PathTextBox.Text.ToLower = "computer" Or PathTextBox.Text.ToLower = "my computer" Then
                                ItemList.Insert(ItemsComboBox.SelectedIndex, NameTextBox.Text & " ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")
                            Else
                                ItemList.Insert(ItemsComboBox.SelectedIndex, NameTextBox.Text & " " & PathTextBox.Text)
                            End If
                        End If
                    ElseIf Not PathTextBox.Text = "" And (System.IO.File.Exists(PathTextBox.Text) = False And System.IO.Directory.Exists(PathTextBox.Text) = False And Not PathTextBox.Text.ToLower = "computer" And Not PathTextBox.Text.ToLower = "my computer") Then
                        MsgBox("The specified file does not exist!")
                        Exit Sub
                    ElseIf PathTextBox.Text = "" Then
                        If ItemList.Count > 0 Then
                            ItemList.Insert(If(Not ItemsComboBox.SelectedIndex = -1, ItemsComboBox.SelectedIndex, 0), " catagory " & NameTextBox.Text)
                        Else
                            ItemList.Add(" catagory " & NameTextBox.Text)
                        End If
                    ElseIf PathTextBox.Text = "" And ItemList.Contains(" catagory " & NameTextBox.Text) = True Then
                        MsgBox("The specified catagory name already exists!")
                        Exit Sub
                    End If
                Else
                    If UseUnknownFileIconCheckBox.IsChecked = True Then
                        ItemList.Insert(ItemsComboBox.SelectedIndex, NameTextBox.Text & " " & PathTextBox.Text & " UseUnknownFileIcon")
                    Else
                        ItemList.Insert(ItemsComboBox.SelectedIndex, NameTextBox.Text & " " & PathTextBox.Text)
                    End If
                End If

            ElseIf Method = "Edit" Then

                If Not PathTextBox.Text.StartsWith("http://") Then
                    If Not PathTextBox.Text = "" And (System.IO.File.Exists(PathTextBox.Text) Or System.IO.Directory.Exists(PathTextBox.Text) Or PathTextBox.Text.ToLower = "computer" Or PathTextBox.Text.ToLower = "my computer") Then
                        If UseUnknownFileIconCheckBox.IsChecked = True Then
                            If PathTextBox.Text.ToLower = "computer" Or PathTextBox.Text.ToLower = "my computer" Then
                                ItemList.Item(ItemsComboBox.SelectedIndex) = NameTextBox.Text & " ::{20D04FE0-3AEA-1069-A2D8-08002B30309D} UseUnknownFileIcon"
                            Else
                                ItemList.Item(ItemsComboBox.SelectedIndex) = NameTextBox.Text & " " & PathTextBox.Text & " UseUnknownFileIcon"
                            End If
                        Else
                            If PathTextBox.Text.ToLower = "computer" Or PathTextBox.Text.ToLower = "my computer" Then
                                ItemList.Item(ItemsComboBox.SelectedIndex) = NameTextBox.Text & " ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
                            Else
                                ItemList.Item(ItemsComboBox.SelectedIndex) = NameTextBox.Text & " " & PathTextBox.Text
                            End If
                        End If
                    ElseIf Not PathTextBox.Text = "" And (System.IO.File.Exists(PathTextBox.Text) = False And System.IO.Directory.Exists(PathTextBox.Text) = False And Not PathTextBox.Text.ToLower = "computer" And Not PathTextBox.Text.ToLower = "my computer") Then
                        MsgBox("The specified file does not exist!")
                        Exit Sub
                    ElseIf PathTextBox.Text = "" And ItemList.Contains(NameTextBox.Text) = False Then
                        ItemList.Item(ItemsComboBox.SelectedIndex) = " catagory " & NameTextBox.Text
                    ElseIf PathTextBox.Text = "" And ItemList.Contains(NameTextBox.Text) = True Then
                        MsgBox("The specified catagory name already exists!")
                        Exit Sub
                    End If
                Else
                    If UseUnknownFileIconCheckBox.IsChecked = True Then
                        ItemList.Item(ItemsComboBox.SelectedIndex) = NameTextBox.Text & " " & PathTextBox.Text & " UseUnknownFileIcon"
                    Else
                        ItemList.Item(ItemsComboBox.SelectedIndex) = NameTextBox.Text & " " & PathTextBox.Text
                    End If
                End If

            ElseIf Method = "Move" Then

                If Not ItemsComboBox.SelectedIndex = MoveComboBox.SelectedIndex Then
                    Dim Backup As String = ItemList.Item(ItemsComboBox.SelectedIndex)
                    ItemList.RemoveAt(ItemsComboBox.SelectedIndex)
                    If MoveComboBox.SelectedIndex >= ItemList.Count - 1 Then
                        ItemList.Add(Backup)
                    Else
                        ItemList.Insert(MoveComboBox.SelectedIndex, Backup)
                    End If
                End If

            ElseIf Method = "Remove" Then
                ItemList.RemoveAt(ItemsComboBox.SelectedIndex)
            End If

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

        Save()
        SetJumpListItems()

        If CloseAfterOKCheckbox.IsChecked = True Then
            End
        Else
            Window_Loaded()
        End If
    End Sub

    Private Sub BrowseButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BrowseButton.Click
        Dim Res As Forms.DialogResult = Browser.ShowDialog()
        If Res = Forms.DialogResult.OK Then
            PathTextBox.Text = Browser.FileName
        End If
    End Sub

    Private Sub CloseAfterOKCheckbox_Checked(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles CloseAfterOKCheckbox.Checked
        SaveSetting("LaunchFromTaskbar2", "Settings", "CloseAfterOKCheckbox", CloseAfterOKCheckbox.IsChecked)
    End Sub

    Private Sub CloseAfterOKCheckbox_Unchecked(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles CloseAfterOKCheckbox.Unchecked
        SaveSetting("LaunchFromTaskbar2", "Settings", "CloseAfterOKCheckbox", CloseAfterOKCheckbox.IsChecked)
    End Sub
End Class
