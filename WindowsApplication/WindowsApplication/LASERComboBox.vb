Public Class LASERComboBox
    Public Event DropDown(sender As Object, e As EventArgs)
    Public Event SelectedIndexChanged(sender As Object, e As EventArgs)
    Private US As New Form
    Private lst As New ListBox

    Public ReadOnly Property LASERCMB As ComboBox
        Get
            Return cmb
        End Get
    End Property

    Private Sub EasySuggestComboBox_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        cmb.Left = 0
        cmb.Top = 0
        cmb.Width = Me.Width
        Me.Height = cmb.Top + cmb.Height
    End Sub

    Private Sub EasySuggestComboBox_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Height = cmb.Height
        Me.BringToFront()
        US.Left = Me.Left
        US.Width = cmb.Width
        US.FormBorderStyle = FormBorderStyle.None
        US.ShowInTaskbar = False
        US.StartPosition = FormStartPosition.Manual
        US.Location = New Point(Me.cmb.Location.X, Me.cmb.Location.Y + Me.cmb.Height)
        lst.Top = 0
        lst.Left = 0
        lst.Width = US.Width
        lst.Height = US.Height
        lst.Font = cmb.Font
        US.Controls.Add(lst)
        US.Show(Me.ParentForm)
        US.Visible = False
        AddHandler lst.MouseClick, AddressOf lst_MouseClick
        AddHandler lst.MouseMove, AddressOf lst_MouseMove
        AddHandler Me.ParentForm.Move, AddressOf ParentForm_Move
    End Sub

    Private Sub ParentForm_Move(sender As Object, e As EventArgs)
        If TypeOf Me.Parent Is GroupBox Then
            US.Location = New Point(Me.ParentForm.Location.X + Me.Parent.Location.X + Me.Location.X + cmb.Location.X + 8, Me.ParentForm.Top + Me.Parent.Top + Me.Top + cmb.Top + cmb.Height + 28)
        Else
            US.Location = New Point(Me.ParentForm.Left + Me.Left + cmb.Left + 8, Me.ParentForm.Top + Me.Top + cmb.Top + cmb.Height + 28)
        End If
    End Sub

    Private Sub cmb_DropDown(sender As Object, e As EventArgs) Handles cmb.DropDown
        RaiseEvent DropDown(sender, e)
        Me.Height = cmb.Top + cmb.Height
    End Sub

    Private Sub cmb_LostFocus(sender As Object, e As EventArgs) Handles cmb.LostFocus
        If ActiveControl Is Nothing Then
            Exit Sub
        End If
        If ActiveControl.TabIndex = 0 Then Exit Sub
        US.Visible = False
    End Sub

    Private Sub lst_MouseClick(sender As Object, e As MouseEventArgs)
        cmb.Text = lst.SelectedItem.ToString
        US.Visible = False
        Me.Height = cmb.Top + cmb.Height
        cmb.Focus()
    End Sub

    Private Sub lst_MouseMove(sender As Object, e As MouseEventArgs)
        Dim point As Point = lst.PointToClient(Cursor.Position)
        Dim index As Integer = lst.IndexFromPoint(point)
        If index < 0 Then Exit Sub
        lst.GetItemRectangle(index).Inflate(1, 2)
        lst.SelectedIndex = index
    End Sub

    Private Sub cmb_KeyDown(sender As Object, e As KeyEventArgs) Handles cmb.KeyDown
        If cmb.DroppedDown = True Then Exit Sub
        If US.Visible = True Then
            If (e.KeyCode = System.Windows.Forms.Keys.Up) Then
                If lst.SelectedIndex > -1 Then
                    lst.SelectedIndex = lst.SelectedIndex - 1
                Else
                    lst.SelectedIndex = (lst.Items.Count - 1)
                End If
                e.Handled = True
            ElseIf (e.KeyCode = System.Windows.Forms.Keys.Down) Then
                If lst.SelectedIndex < (lst.Items.Count - 1) Then
                    lst.SelectedIndex = lst.SelectedIndex + 1
                Else
                    lst.SelectedIndex = -1
                End If
                e.Handled = True
            ElseIf (e.KeyCode = System.Windows.Forms.Keys.Enter) Then
                e.SuppressKeyPress = True
                If US.Visible = True And lst.SelectedIndex > -1 And lst.SelectedIndex < (lst.Items.Count) Then
                    cmb.Text = lst.Items.Item(lst.SelectedIndex)
                    US.Visible = False
                End If
                cmb.SelectAll()
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub cmb_Leave(sender As Object, e As EventArgs) Handles cmb.Leave
        If ActiveControl Is Nothing Then
            If US.Visible = True And lst.SelectedIndex > -1 Then cmb.Text = lst.Items.Item(lst.SelectedIndex)
            US.Visible = False
        ElseIf ActiveControl.Name <> lst.Name Then
            If US.Visible = True And lst.SelectedIndex > -1 Then cmb.Text = lst.Items.Item(lst.SelectedIndex)
            US.Visible = False
        End If
    End Sub

    Private Sub cmb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb.SelectedIndexChanged
        RaiseEvent SelectedIndexChanged(sender, e)
    End Sub

    Private Sub cmb_TextUpdate(sender As Object, e As EventArgs) Handles cmb.TextUpdate
        Dim textToSearch As String = cmb.Text.ToLower()
        Me.Height = cmb.Top + cmb.Height
        If String.IsNullOrEmpty(textToSearch) Then Exit Sub
        Dim result As String() = (From i As String In cmb.Items Where i.ToLower().Contains(textToSearch) Select i).ToArray()
        If result.Length = 0 Then Exit Sub
        lst.Items.Clear()
        lst.Items.AddRange(result)
        If TypeOf Me.Parent Is GroupBox Then
            US.Location = New Point(Me.ParentForm.Location.X + Me.Parent.Location.X + Me.Location.X + cmb.Location.X + 8, Me.ParentForm.Top + Me.Parent.Top + Me.Top + cmb.Top + cmb.Height + 28)
        Else
            US.Location = New Point(Me.ParentForm.Left + Me.Left + cmb.Left + 8, Me.ParentForm.Top + Me.Top + cmb.Top + cmb.Height + 28)
        End If
        US.Visible = True
        cmb.Focus()
        cmb.Select(cmb.Text.Length, 0)
        If ((lst.Items.Count + 1) * lst.ItemHeight) < (21 * lst.ItemHeight) Then
            lst.Height = (lst.Items.Count + 1) * lst.ItemHeight
        Else
            lst.Height = (21) * lst.ItemHeight
        End If
        US.Height = lst.Top + lst.Height
    End Sub
End Class
