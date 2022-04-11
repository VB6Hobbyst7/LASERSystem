Public Class frmSupplier

    Private Sub cmdNew_Click(sender As Object, e As EventArgs) Handles cmdNew.Click
        CMD = New OleDb.OleDbCommand("SELECT top 1 SuNo from Supplier ORDER BY SuNo Desc;", CNN)
        DR = CMD.ExecuteReader(CommandBehavior.CloseConnection)
        If DR.HasRows = True Then
            DR.Read()
            txtSuNo.Text = Int(DR.Item("SuNo")) + 1
        End If
        cmbSuName.Text = ""
        txtSuTelNo.Text = ""
        cmdSave.Text = "Save"
        cmdDelete.Enabled = False
    End Sub

    Private Sub frmSupplier_Leave(sender As Object, e As EventArgs) Handles Me.Leave
        Me.Close()
    End Sub

    Private Sub frmSupplier_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetCNN()
        cmbFilter.Items.Clear()     'add values of cmdFilters
        cmbFilter.Items.Add("by Supplier No")
        cmbFilter.Items.Add("by Supplier Name")
        cmbFilter.Items.Add("by Supplier Telephone No")
        cmbFilter.Items.Add("by All")
        cmbFilter.Text = "by All"
        txtSearch.Text = ""
        Call txtSearch_TextChanged(sender, e)   'refresh grdstock
        Call cmdNew_Click(sender, e)
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Call frmSupplier_Leave(sender, e)
    End Sub

    Private Sub cmbSuName_DropDown(sender As Object, e As EventArgs) Handles cmbSuName.DropDown
        CMD = New OleDb.OleDbCommand("Select SuName from Supplier group by  SuName;", CNN)
        DR = CMD.ExecuteReader()
        If DR.HasRows = True Then
            cmbSuName.Items.Clear()
            While DR.Read
                cmbSuName.Items.Add(DR("SuName").ToString)
            End While
        Else
            cmbSuName.Items.Clear()
        End If
    End Sub

    Private Sub cmbSuName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSuName.SelectedIndexChanged
        CMD = New OleDb.OleDbCommand("SELECT * from Supplier where SuName='" & cmbSuName.Text & "';", CNN)
        DR = CMD.ExecuteReader()
        If DR.HasRows = True Then
            DR.Read()
            txtSuNo.Text = DR("SuNo").ToString
            cmbSuName.Text = DR("SuName").ToString
            txtSuTelNo.Text = DR("SuTelNo").ToString
            For Each Row As DataGridViewRow In grdSupplier.Rows
                If Row.Cells(0).Value.ToString = txtSuNo.Text Then
                    Row.Selected = True
                    grdSupplier.Select()
                    Exit Sub
                End If
            Next Row
            cmdSave.Text = "Edit"
            cmdDelete.Enabled = True
        End If
    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        If txtSuNo.Text = "" Then
            MsgBox("Please enter supplier no", vbOKOnly + vbExclamation)
            txtSuNo.Focus()
            Exit Sub
        ElseIf cmbSuName.Text = "" Then
            MsgBox("Please enter supplier name", vbOKOnly + vbExclamation)
            cmbSuName.Focus()
            Exit Sub
        End If
        Select Case cmdSave.Text
            Case "Save"
                CMD = New OleDb.OleDbCommand("Select SuNo from Supplier where SuNo =" & txtSuNo.Text & ";", CNN)
                DR = CMD.ExecuteReader()
                If DR.HasRows = True Then
                    MsgBox("Supplier No is exist", vbOKOnly + vbExclamation)
                    txtSuNo.Focus()
                    Exit Sub
                End If
                CMD = New OleDb.OleDbCommand("Select SuName from Supplier where SuName ='" & cmbSuName.Text & "';", CNN)
                DR = CMD.ExecuteReader()
                If DR.HasRows = True Then
                    MsgBox("Supplier Name is exist", vbOKOnly + vbExclamation)
                    cmbSuName.Focus()
                    Exit Sub
                End If
                CMD = New OleDb.OleDbCommand("Insert into Supplier(SuNo,SuName,SuTelNo) Values(" & txtSuNo.Text & ",'" & cmbSuName.Text & "','" & txtSuTelNo.Text & "');", CNN)
                CMD.ExecuteNonQuery()
                Call txtSearch_TextChanged(sender, e)
                cmdSave.Text = "Edit"
                cmdDelete.Enabled = True
                MsgBox("Save Successful", vbExclamation + vbOKOnly)
            Case "Edit"
                If MsgBox("Are you sure edit?", vbYesNo + vbInformation) = vbYes Then
                    CMD = New OleDb.OleDbCommand("Update Supplier set SuNo=" & txtSuNo.Text & _
                                                 ",SuName = '" & cmbSuName.Text & "'" & _
                                                 ",SuTelNo =  '" & txtSuTelNo.Text & "'" & _
                                                 " where SuNo=" & txtSuNo.Text, CNN)
                    CMD.ExecuteNonQuery()
                    Call txtSearch_TextChanged(sender, e)
                End If
        End Select
        For Each Row As DataGridViewRow In grdSupplier.Rows
            If Row.Cells(0).Value.ToString = txtSuNo.Text Then
                Row.Selected = True
                grdSupplier.Select()
                Exit Sub
            End If
        Next Row
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        If MsgBox("Are you sure delete?", vbYesNo + vbInformation) = vbYes Then
            CMD = New OleDb.OleDbCommand("DELETE from Supplier where SuNo=" & txtSuNo.Text, CNN)
            CMD.ExecuteNonQuery()
            Call txtSearch_TextChanged(sender, e)
            Call cmdNew_Click(sender, e)
        End If
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        Dim dt As New DataTable
        Dim x As String = ""
        grdSupplier.ClearSelection()
        If txtSearch.Text <> "" Then
            Select Case cmbFilter.Text
                Case "by Supplier No"
                    x = "Where SuNo Like '%" & txtSearch.Text & "%'"
                Case "by Supplier Name"
                    x = "Where SuName like '%" & txtSearch.Text & "%'"
                Case "by Supplier Telephone No"
                    x = "Where SuTelNo like '%" & txtSearch.Text & "%'"
                Case "by All"
                    x = "Where SuNo like '%" & txtSearch.Text & "%' or SuName like '%" & txtSearch.Text & "%' or SuTelNo like '%" & txtSearch.Text & "%'"
            End Select
        Else
            x = "Order by SuNo"
        End If
        Dim da As New OleDb.OleDbDataAdapter("SELECT SuNo as [No],SuName as [Name],SuTelNo as [Telephone No] from Supplier " & x & ";", CNN)
        da.Fill(dt)
        Me.grdSupplier.DataSource = dt
        grdSupplier.Refresh()
    End Sub

    Private Sub grdSupplier_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdSupplier.CellContentClick

    End Sub

    Private Sub grdSupplier_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdSupplier.CellContentDoubleClick
        Dim index As Integer
        index = e.RowIndex
        Dim selectedrow As DataGridViewRow
        If index >= 0 Then
            selectedrow = grdSupplier.Rows(index)
            cmbSuName.Text = selectedrow.Cells(1).Value.ToString
            Call cmbSuName_SelectedIndexChanged(sender, e)
            cmdSave.Text = "Edit"   'Change edit mode
            cmdDelete.Enabled = True
        End If
    End Sub
End Class