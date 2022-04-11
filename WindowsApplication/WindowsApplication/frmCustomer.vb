Public Class frmCustomer
    Public Property Caller As String = ""

    Private Sub frmCustomer_Leave(sender As Object, e As EventArgs) Handles Me.Leave
        Me.Close()
        Me.Tag = ""
    End Sub

    Private Sub cmbCuName_DropDown(sender As Object, e As EventArgs) Handles cmbCuName.DropDown
        CMD = New OleDb.OleDbCommand("Select CuName from Customer group by  CuName;", CNN)
        DR = CMD.ExecuteReader()
        If DR.HasRows = True Then
            cmbCuName.Items.Clear()
            While DR.Read
                cmbCuName.Items.Add(DR("CuName").ToString)
            End While
        Else
            cmbCuName.Items.Clear()
        End If
    End Sub

    Private Sub frmCustomer_Load(sender As Object, e As EventArgs) Handles Me.Load
        GetCNN()
        MenuStrip.Items.Add(mnustrpMENU)
        cmbFilter.Items.Clear()     'add values of cmdFilters
        cmbFilter.Items.Add("by Customer No")
        cmbFilter.Items.Add("by Customer Name")
        cmbFilter.Items.Add("by Customer Telephone No 1")
        cmbFilter.Items.Add("by Customer Telephone No 2")
        cmbFilter.Items.Add("by Customer Telephone No 3")
        cmbFilter.Items.Add("by All")
        cmbFilter.Text = "by All"
        txtSearch.Text = ""
        Call txtSearch_TextChanged(sender, e)   'refresh grdstock
        Call cmdNew_Click(sender, e)
        Call cmbCuName_DropDown(sender, e)
        If Me.Tag = "" Then
            cmdDone.Enabled = False
        Else
            cmdDone.Enabled = True
        End If
        Select Case Me.Tag
            Case "Repair"
                cmbCuName.Text = frmRepair.cmbCuName.Text
                Call cmbCuName_SelectedIndexChanged(sender, e)
        End Select
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        Dim dt As New DataTable
        Dim x As String = ""
        If txtSearch.Text <> "" Then
            Select Case cmbFilter.Text
                Case "by Customer No"
                    x = "Where CuNo Like '%" & txtSearch.Text & "%'"
                Case "by Customer Name"
                    x = "Where CuName like '%" & txtSearch.Text & "%'"
                Case "by Customer Telephone No 1"
                    x = "Where CuTelNo1 like '%" & txtSearch.Text & "%'"
                Case "by Customer Telephone No 2"
                    x = "Where CuTelNo2 like '%" & txtSearch.Text & "%'"
                Case "by Customer Telephone No 3"
                    x = "Where CuTelNo3 like '%" & txtSearch.Text & "%'"
                Case "by All"
                    x = "Where CuNo like '%" & txtSearch.Text & "%' or CuName like '%" & txtSearch.Text & "%' or CuTelNo1 like '%" & txtSearch.Text & "%' or CuTelNo2 like '%" & txtSearch.Text & "%' or CuTelNo3 like '%" & txtSearch.Text & "%'"
            End Select
        Else
            x = "Order by CuNo"
        End If
        Dim da As New OleDb.OleDbDataAdapter("SELECT CuNo as [No],CuName as [Name],CuTelNo1 as [Telephone No 1],CuTelNo2 as [Telephone No 2],CuTelNo3 as [Telephone No 3] from Customer " & x & ";", CNN)
        da.Fill(dt)
        Me.grdCustomer.DataSource = dt
        grdCustomer.Refresh()
    End Sub

    Public Sub cmbCuName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCuName.SelectedIndexChanged
        CMD = New OleDb.OleDbCommand("SELECT * from Customer where CuName='" & cmbCuName.Text & "';", CNN)
        DR = CMD.ExecuteReader()
        If DR.HasRows = True Then
            DR.Read()
            txtCuNo.Text = DR("CuNo").ToString
            cmbCuName.Text = DR("CuName").ToString
            txtCuTelNo1.Text = DR("CuTelNo1").ToString
            txtCuTelNo2.Text = DR("CuTelNo2").ToString
            txtCuTelNo3.Text = DR("CuTelNo3").ToString
            For Each Row As DataGridViewRow In grdCustomer.Rows
                If Row.Cells(0).Value.ToString = txtCuNo.Text Then
                    Row.Selected = True
                    grdCustomer.Select()
                    Exit For
                End If
            Next Row
            cmdSave.Text = "Edit"
            SaveToolStripMenuItem.Text = cmdSave.Text
            cmdDone.Text = "Done"
            DoneSaveToolStripMenuItem.Text = "Done"
            cmdDelete.Enabled = True
            DeleteToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub grdCustomer_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles grdCustomer.CellDoubleClick
        Dim index As Integer
        index = e.RowIndex
        Dim selectedrow As DataGridViewRow
        If index >= 0 Then
            selectedrow = grdCustomer.Rows(index)
            cmbCuName.Text = selectedrow.Cells(1).Value.ToString
            Call cmbCuName_SelectedIndexChanged(sender, e)
        End If
        cmdSave.Text = "Edit"
        SaveToolStripMenuItem.Text = cmdSave.Text
        cmdDone.Text = "Done"
        DoneSaveToolStripMenuItem.Text = cmdDone.Text
        cmdDelete.Enabled = True
        DeleteToolStripMenuItem.Enabled = True
    End Sub

    Private Sub cmdDone_Click(sender As Object, e As EventArgs) Handles cmdDone.Click
        Call cmdSave_Click(sender, e)
        If cmdDone.Tag = "0" Then
            Exit Sub
        End If
        Select Case Me.Tag
            Case "Sale"
                For Each oForm As frmSale In Application.OpenForms().OfType(Of frmSale)()
                    If oForm.Name = Me.Caller Then
                        With oForm
                            .cmbCuName.Text = Me.cmbCuName.Text
                            .txtCuTelNo1.Text = Me.txtCuTelNo1.Text
                            .txtCuTelNo2.Text = Me.txtCuTelNo2.Text
                            .txtCuTelNo3.Text = Me.txtCuTelNo3.Text
                            .txtCuTelNo1.Tag = ""
                        End With
                        Exit For
                    End If
                Next
            Case "Receive"
                For Each oForm As frmReceive In Application.OpenForms().OfType(Of frmReceive)()
                    If oForm.Name = Me.Caller Then
                        With oForm
                            .txtCuTelNo1.Text = Me.txtCuTelNo1.Text
                            .txtCuTelNo2.Text = Me.txtCuTelNo2.Text
                            .txtCuTelNo3.Text = Me.txtCuTelNo3.Text
                            frmReceive.cmbCuName_Text(Me.cmbCuName.Text)
                        End With
                        Exit For
                    End If
                Next
            Case "Repair"
                With frmRepair
                    .cmbCuName.Text = Me.cmbCuName.Text
                    Call .CmbCuName_SelectedIndexChanged(sender, e)
                End With
        End Select
        Me.Tag = ""
        cmdDone.Tag = ""
        Me.Close()
    End Sub

    Private Sub cmdClose_Click(sender As Object, e As EventArgs) Handles cmdClose.Click
        Call frmCustomer_Leave(sender, e)
    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        cmdDone.Tag = ""
        If txtCuNo.Text = "" Then
            MsgBox("Please enter customer no", vbOKOnly + vbExclamation)
            txtCuNo.Focus()
            cmdDone.Tag = "0"
            Exit Sub
        ElseIf cmbCuName.Text = "" Then
            MsgBox("Please enter customer name", vbOKOnly + vbExclamation)
            cmbCuName.Focus()
            cmdDone.Tag = "0"
            Exit Sub
        End If
        Select Case cmdSave.Text
            Case "Save"
                If CheckExistDatacmb(cmbCuName, "Select CuName from Customer where CuName ='" & cmbCuName.Text & "';", "Customer Name is exist", True) = True Then
                    For i As Integer = 0 To 1000
                        CMD = New OleDb.OleDbCommand("Select CuName from Customer Where CuName = '" & cmbCuName.Text & " " & i.ToString & "'", CNN)
                        DR = CMD.ExecuteReader
                        If DR.HasRows = False Then
                            cmbCuName.Text = cmbCuName.Text + " " + i.ToString
                            Exit For
                        End If
                    Next
                    cmdDone.Tag = "0"
                    Exit Sub
                End If
                If CheckExistDatatxt(txtCuNo, "Select CuNo from Customer where CuNo =" & txtCuNo.Text & ";", "Customer No is exist", True) = True Then
                    cmdDone.Tag = "0"
                    Exit Sub
                End If
                CMDUPDATE("Insert into Customer(CuNo,CuName,CuTelNo1,CuTelNo2,CuTelNo3) Values(" & txtCuNo.Text & ",'" & cmbCuName.Text & "','" & txtCuTelNo1.Text & "','" & txtCuTelNo2.Text & "','" & txtCuTelNo3.Text & "');")
                Call txtSearch_TextChanged(sender, e)
                cmdSave.Text = "Edit"
                SaveToolStripMenuItem.Text = cmdSave.Text
                cmdDelete.Enabled = True
                DeleteToolStripMenuItem.Enabled = True
                MsgBox("Save Successful", vbExclamation + vbOKOnly)
            Case "Edit"
                CMDUPDATE("Update Customer set CuNo=" & txtCuNo.Text &
                                                 ",CuName = '" & cmbCuName.Text & "'" &
                                                 ",CuTelNo1 =  '" & txtCuTelNo1.Text & "'" &
                                                 ",CuTelNo2 =  '" & txtCuTelNo2.Text & "'" &
                                                 ",CuTelNo3 =  '" & txtCuTelNo3.Text & "'" &
                                                 " where CuNo=" & txtCuNo.Text)
                Call txtSearch_TextChanged(sender, e)
        End Select
        For Each Row As DataGridViewRow In grdCustomer.Rows
            If Row.Cells(0).Value.ToString = txtCuNo.Text Then
                Row.Selected = True
                grdCustomer.Select()
                Exit For
            End If
        Next Row
    End Sub

    Private Sub cmdNew_Click(sender As Object, e As EventArgs) Handles cmdNew.Click
        Call AutomaticPrimaryKey(txtCuNo, "SELECT top 1 CuNo from Customer ORDER BY CuNo Desc;", "CuNo")
        cmbCuName.Text = ""
        txtCuTelNo1.Text = ""
        txtCuTelNo2.Text = ""
        txtCuTelNo3.Text = ""
        cmdSave.Text = "Save"
        SaveToolStripMenuItem.Text = cmdSave.Text
        cmdDone.Text = "Done + Save"
        DoneSaveToolStripMenuItem.Text = cmdDone.Text
        cmdDelete.Enabled = False
        DeleteToolStripMenuItem.Enabled = False
    End Sub

    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        Dim C As Integer = 0
        Dim tmp As String = ""
        CMD = New OleDb.OleDbCommand("Select * from Sale where CuNo= " & txtCuNo.Text & ";", CNN)
        DR = CMD.ExecuteReader()
        While DR.Read
            C = C + 1
            tmp = tmp + "Sale: " + DR("SaNo").ToString + vbCrLf
        End While
        CMD = New OleDb.OleDbCommand("Select * from Receive where CuNo= " & txtCuNo.Text & ";", CNN)
        DR = CMD.ExecuteReader()
        While DR.Read
            C = C + 1
            tmp = tmp + "Receive: " + DR("RNo").ToString + vbCrLf
        End While
        If C > 0 Then
            MsgBox("Relationship/s " & C & " ක් සොයා ගැනුනි.ඒ නිසා මෙම Customer ව Delete කිරීමට හැකියාවක් නොමැත නමුත්, ඔබට එම Relationship/s ඉවත් කිරිමට හැකිනම්, ඒ සඳහා ඉඩ ලබා දෙන්නෙම්. ඒවා, " + vbCrLf + tmp, vbCritical + vbOKOnly)
            Exit Sub
        End If
        If MsgBox("Are you sure delete?", vbYesNo + vbInformation) = vbYes Then
            CMDUPDATE("DELETE from Customer where CuNo=" & txtCuNo.Text)
            Call txtSearch_TextChanged(sender, e)
            Call cmdNew_Click(sender, e)
        End If
    End Sub

    Private Sub txtCuTelNo1_KeyUp(sender As Object, e As KeyEventArgs) Handles txtCuTelNo1.KeyUp
        If txtCuTelNo1.Text <> "          " Then
            CMD = New OleDb.OleDbCommand("Select * from Customer where CuTelNo1='" & txtCuTelNo1.Text & "' or CuTelNo2='" & txtCuTelNo1.Text & "' or CuTelNo3='" & txtCuTelNo1.Text & "';", CNN)
            DR = CMD.ExecuteReader()
            If DR.HasRows = True Then
                If MsgBox("Another Customer was found assigned this Telephone No. Will it be opened?", vbYesNo + vbCritical) = vbYes Then
                    DR.Read()
                    cmbCuName.Text = DR("CuName").ToString
                    Call cmbCuName_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub txtCuTelNo2_KeyUp(sender As Object, e As KeyEventArgs) Handles txtCuTelNo2.KeyUp
        If txtCuTelNo2.Text <> "          " Then
            CMD = New OleDb.OleDbCommand("Select * from Customer where CuTelNo1='" & txtCuTelNo2.Text & "' or CuTelNo2='" & txtCuTelNo2.Text & "' or CuTelNo3='" & txtCuTelNo2.Text & "';", CNN)
            DR = CMD.ExecuteReader()
            If DR.HasRows = True Then
                If MsgBox("Another Customer was found assigned this Telephone No. Will it be opened?", vbYesNo + vbCritical) = vbYes Then
                    DR.Read()
                    cmbCuName.Text = DR("CuName").ToString
                    Call cmbCuName_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub txtCuTelNo3_KeyUp(sender As Object, e As KeyEventArgs) Handles txtCuTelNo3.KeyUp
        If txtCuTelNo3.Text <> "          " Then
            CMD = New OleDb.OleDbCommand("Select * from Customer where CuTelNo1='" & txtCuTelNo3.Text & "' or CuTelNo2='" & txtCuTelNo3.Text & "' or CuTelNo3='" & txtCuTelNo3.Text & "';", CNN)
            DR = CMD.ExecuteReader()
            If DR.HasRows = True Then
                If MsgBox("Another Customer was found assigned this Telephone No. Will it be opened?", vbYesNo + vbCritical) = vbYes Then
                    DR.Read()
                    cmbCuName.Text = DR("CuName").ToString
                    Call cmbCuName_SelectedIndexChanged(sender, e)
                End If
            End If
        End If
    End Sub

    Private Sub DoneSaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DoneSaveToolStripMenuItem.Click
        Call cmdDone_Click(sender, e)
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        Call cmdNew_Click(sender, e)
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        Call cmdSave_Click(sender, e)
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Call cmdDelete_Click(sender, e)
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Call cmdClose_Click(sender, e)
    End Sub
End Class