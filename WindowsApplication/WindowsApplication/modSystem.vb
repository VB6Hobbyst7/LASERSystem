Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Net
Imports System.Text
Imports System.Web.Security
Imports Guna.UI2.WinForms
Imports System.IO
Module modSystem
    Public CNN As New OleDb.OleDbConnection
    Public CMD As New OleDb.OleDbCommand
    Public DR As OleDb.OleDbDataReader
    Public DA As OleDb.OleDbDataAdapter
    Public DT As DataTable
    Public DS As DataSet

    Public Sub GetCNN()
        Try
            Dim Simple As New Simple3Des("RandomKey45")
            CNN = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & My.Settings.DatabaseCNN &
                                      ";Jet OLEDB:Database Password=" & Simple.Decode(My.Settings.DBPassword) & ";")
            CNN.Open()
        Catch ex As Exception
            MsgBox("Message:   " + ex.Message + vbCrLf + "Please inform this error to developer for fixing",
                   vbCritical + vbOKOnly, "LASER System - Database Issue")

        End Try
    End Sub

    Public Sub CmbDropDown(cmb As ComboBox, SQL As String, ColumnName As String)
        Dim CMD0 = New OleDbCommand(SQL, CNN)
        Dim DR0 As OleDbDataReader = CMD0.ExecuteReader()
        cmb.Items.Clear()
        cmb.AutoCompleteCustomSource.Clear()
        If DR0.HasRows = True Then
            While DR0.Read
                cmb.Items.Add(DR0(ColumnName).ToString)
                cmb.AutoCompleteCustomSource.Add(DR0(ColumnName).ToString)
            End While
        End If
        CMD0.Cancel()
        DR0.Close()
    End Sub

    Public Sub AutomaticPrimaryKey(txt As TextBox, SQL As String, ColumnName As String)
        Dim CMD0 = New OleDbCommand(SQL, CNN)
        Dim DR0 As OleDbDataReader = CMD0.ExecuteReader()
        If DR0.HasRows = True Then
            DR0.Read()
            txt.Text = Int(DR0.Item(ColumnName)) + 1
        Else
            txt.Text = "1"
        End If
        CMD0.Cancel()
        DR0.Close()
    End Sub

    Public Function AutomaticPrimaryKeyStr(TableName As String, ColumnName As String) As String
        Dim CMD0 As OleDbCommand = New OleDbCommand("Select Top 1 " & ColumnName & " from " & TableName & " Order by " & ColumnName & " Desc", CNN)
        Dim tmp As String = CMD0.ExecuteScalar()
        If tmp <> "" Then
            Return (Int(tmp) + 1)
        Else
            Return ("1")
        End If
        CMD0.Cancel()
    End Function

    Public Function GetRowsCount(Cmd As OleDbCommand) As Integer
        Dim DR0 As OleDbDataReader = Cmd.ExecuteReader
        Dim DT0 As New DataTable
        DT0.Load(DR0)
        Return (DT0.Rows.Count)
        DR0.Close()
        DT0.Clear()
    End Function

    Public Sub OnlynumberQty(e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 46 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Public Sub OnlynumberPrice(e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 46 And Asc(e.KeyChar) <> 45 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Public Sub TextBoxPrice_keyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        OnlynumberPrice(e)
    End Sub

    Public Sub TextBoxQty_keyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        OnlynumberQty(e)
    End Sub

    Public Function GetStrfromRelatedfield(SQL As String, ColumnName As String) As String
        Dim CMD0 As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQL, CNN)
        If IsDBNull(CMD0.ExecuteScalar) = False Then
            Return CMD0.ExecuteScalar
        Else
            Return ""
        End If
    End Function

    Public Function CheckEmptyStr(str As String, msg As String) As Boolean
        CheckEmptyStr = True
        If str = "" Then
            MsgBox(msg, vbCritical + vbOKOnly)
            Return False
        End If
    End Function

    Public Function CheckEmptyfieldtxt(txt As TextBox, msg As String) As Boolean
        CheckEmptyfieldtxt = True
        If txt.Text = "" Then
            MsgBox(msg, vbCritical + vbOKOnly)
            txt.Focus()
            Return False
        End If
    End Function

    Public Function CheckEmptyfieldcmb(cmb As ComboBox, msg As String) As Boolean
        CheckEmptyfieldcmb = True
        If cmb.Text = "" Then
            MsgBox(msg, vbCritical + vbOKOnly)
            cmb.Focus()
            Return False
        End If
    End Function

    Public Function CheckExistData(SQL As String) As Boolean
        Dim CMD0 = New OleDbCommand(SQL, CNN)
        Dim DR0 As OleDbDataReader = CMD0.ExecuteReader()
        If DR0.HasRows = True Then
            Return True
        Else
            Return False
        End If
        CMD0.Cancel()
        DR0.Close()
    End Function

    Public Function CheckExistDatacmb(cmb As ComboBox, SQL As String, msg As String, IsDataExist As Boolean) As Boolean
        Dim CMD0 = New OleDbCommand(SQL, CNN)
        Dim DR0 As OleDbDataReader = CMD0.ExecuteReader()
        If DR0.HasRows = True Then
            If IsDataExist = True Then
                MsgBox(msg, vbCritical + vbOKOnly)
                cmb.Focus()
            End If
            Return True
        Else
            If IsDataExist = False Then
                MsgBox(msg, vbCritical + vbOKOnly)
                cmb.Focus()
            End If
            Return False
        End If
        CMD0.Cancel()
        DR0.Close()
    End Function

    Public Function CheckExistDatatxt(txt As TextBox, SQL As String, msg As String, IsDataExist As Boolean) As Boolean
        Dim CMD0 = New OleDbCommand(SQL, CNN)
        Dim DR0 As OleDbDataReader = CMD0.ExecuteReader()
        If DR0.HasRows = True Then
            If IsDataExist = True Then
                MsgBox(msg, vbCritical + vbOKOnly)
                txt.Focus()
            End If
            Return True
        Else
            If IsDataExist = False Then
                MsgBox(msg, vbCritical + vbOKOnly)
                txt.Focus()
            End If
            Return False
        End If
        CMD0.Cancel()
        DR0.Close()
    End Function

    Public Sub WriteActivity(txt As String)
        File.AppendAllText(Application.StartupPath & "\System Files\Activity\Activity.ls", txt + vbNewLine)
        MdifrmMain.txtActivity.Text = ""
        MdifrmMain.txtActivity.Text = File.ReadAllText(Application.StartupPath & "\System Files\Activity\Activity.ls")
    End Sub

    Public Function CheckExistRelationsforDelete(SQl As String, FieldName As String, msg As String) As Boolean
        CheckExistRelationsforDelete = True
        CMD = New OleDb.OleDbCommand(SQl, CNN)
        DR = CMD.ExecuteReader
        Dim tmp As String = ""
        If DR.HasRows = True Then
            While DR.Read
                tmp = tmp + DR(FieldName).ToString + " "
            End While
            MsgBox(msg + vbCrLf + "The field/s called '" + FieldName + "' are/is " + tmp, vbCritical + vbOKOnly)
            Return False
        End If
    End Function

    Public Sub CMDUPDATE(str As String, Optional ByVal DoesAdminSend As Boolean = False, Optional ByVal Remarks As String = "")
        Dim CMDUPDATEDB As OleDbCommand
        If DoesAdminSend = True Then
            CMDUPDATEDB = New OleDbCommand("Insert into AdminPermission(APNo,APDate,Status,Command,AppliedUNo,Remarks) Values(" &
                                           AutomaticPrimaryKeyStr("AdminPermission", "APNo") &
                                           ",#" & DateAndTime.Now.ToString &
                                           "#,'Waiting',""" & str.ToString.Replace("""", """""") & """," & MdifrmMain.Tag & ",'" & Remarks & "')", CNN)
            CMDUPDATEDB.ExecuteNonQuery()
            CMDUPDATEDB = New OleDbCommand("Insert into OnlineDB(ID,ODate,Command) Values(" &
                                           AutomaticPrimaryKeyStr("OnlineDB", "ID") & ",#" & DateAndTime.Now &
                                           "#,""" &
                                           "Insert into AdminPermission(APNo, APDate, Status, Command, AppliedUNo, Remarks) Values(" &
                                           AutomaticPrimaryKeyStr("AdminPermission", "APNo") &
                                           ",#" & DateAndTime.Now.ToString &
                                           "#,'Waiting',""""" & str.ToString.Replace("""", """""") & """""," & MdifrmMain.Tag & ",'" & Remarks & "')" &
                                           """)", CNN)
            CMDUPDATEDB.ExecuteNonQuery()
        Else
            If str.Contains("?") = True Then
                Dim splittxt() As String = str.Split("?")
                Select Case splittxt(1)
                    Case "PrimaryKey"
                        str = str.Replace("?" + splittxt(1) + "?" + splittxt(2) + "?" + splittxt(3) + "?", AutomaticPrimaryKeyStr(splittxt(2), splittxt(3)))
                End Select
            End If
            CMDUPDATEDB = New OleDb.OleDbCommand(str, CNN)
            CMDUPDATEDB.ExecuteNonQuery()
            CMDUPDATEDB = New OleDbCommand("Insert into OnlineDB(ID,ODate,Command) Values(" &
                                           AutomaticPrimaryKeyStr("OnlineDB", "ID") & ",#" & DateAndTime.Now & "#,""" &
                                           str.ToString.Replace("""", """""") & """)", CNN)
            CMDUPDATEDB.ExecuteNonQuery()
            WriteActivity(DateAndTime.Now + "- " + str)
        End If
        CMDUPDATEDB.Cancel()
    End Sub

    Public Function NextfrmNo(frmNew As Form) As Integer
        Dim i As Integer = 0
        For Each oForm As Form In Application.OpenForms().OfType(Of Form)()
            If oForm.Name.ToString.StartsWith(frmNew.Name) = True Then
                If IsNumeric(oForm.Name.ToString.Replace(frmNew.Name, "")) = True Then
                    If i <= oForm.Name.ToString.Replace(frmNew.Name, "") Then
                        i = oForm.Name.ToString.Replace(frmNew.Name, "")
                    End If
                End If
            End If
        Next
        Return (i + 1)
    End Function

#Region "OpenCashDrawer"
    Public Sub OpenCashdrawer()
        'Modify DrawerCode to your receipt printer open drawer code
        Dim DrawerCode As String = Chr(27) & Chr(112) & Chr(48) & Chr(64) & Chr(64)
        'Modify PrinterName to your receipt printer name
        Dim PrinterName As String = "Zonerich AB-220K"

        RawPrinter.PrintRaw(PrinterName, DrawerCode)
    End Sub

    Public Class RawPrinter
        ' ----- Define the data type that supplies basic print job information to the spooler.
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Public Structure DOCINFO
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pDocName As String
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pOutputFile As String
            <MarshalAs(UnmanagedType.LPWStr)>
            Public pDataType As String
        End Structure

        ' ----- Define interfaces to the functions supplied in the DLL.
        <DllImport("winspool.drv", EntryPoint:="OpenPrinterW", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
        Public Shared Function OpenPrinter(ByVal printerName As String, ByRef hPrinter As IntPtr, ByVal printerDefaults As Integer) As Boolean
        End Function

        <DllImport("winspool.drv", EntryPoint:="ClosePrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
        Public Shared Function ClosePrinter(ByVal hPrinter As IntPtr) As Boolean
        End Function

        <DllImport("winspool.drv", EntryPoint:="StartDocPrinterW", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
        Public Shared Function StartDocPrinter(ByVal hPrinter As IntPtr, ByVal level As Integer, ByRef documentInfo As DOCINFO) As Boolean
        End Function

        <DllImport("winspool.drv", EntryPoint:="EndDocPrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
        Public Shared Function EndDocPrinter(ByVal hPrinter As IntPtr) As Boolean
        End Function

        <DllImport("winspool.drv", EntryPoint:="StartPagePrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
        Public Shared Function StartPagePrinter(ByVal hPrinter As IntPtr) As Boolean
        End Function

        <DllImport("winspool.drv", EntryPoint:="EndPagePrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
        Public Shared Function EndPagePrinter(ByVal hPrinter As IntPtr) As Boolean
        End Function

        <DllImport("winspool.drv", EntryPoint:="WritePrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
        Public Shared Function WritePrinter(ByVal hPrinter As IntPtr, ByVal buffer As IntPtr, ByVal bufferLength As Integer, ByRef bytesWritten As Integer) As Boolean
        End Function

        Public Shared Function PrintRaw(ByVal printerName As String, ByVal origString As String) As Boolean
            ' ----- Send a string of  raw data to  the printer.
            Dim hPrinter As IntPtr
            Dim spoolData As New DOCINFO
            Dim dataToSend As IntPtr
            Dim dataSize As Integer
            Dim bytesWritten As Integer

            ' ----- The internal format of a .NET String is just
            '       different enough from what the printer expects
            '       that there will be a problem if we send it
            '       directly. Convert it to ANSI format before
            '       sending.
            dataSize = origString.Length()
            dataToSend = Marshal.StringToCoTaskMemAnsi(origString)

            ' ----- Prepare information for the spooler.
            spoolData.pDocName = "OpenDrawer" ' class='highlight'
            spoolData.pDataType = "RAW"

            Try
                ' ----- Open a channel to  the printer or spooler.
                Call OpenPrinter(printerName, hPrinter, 0)

                ' ----- Start a new document and Section 1.1.
                Call StartDocPrinter(hPrinter, 1, spoolData)
                Call StartPagePrinter(hPrinter)

                ' ----- Send the data to the printer.
                Call WritePrinter(hPrinter, dataToSend,
                   dataSize, bytesWritten)

                ' ----- Close everything that we opened.
                EndPagePrinter(hPrinter)
                EndDocPrinter(hPrinter)
                ClosePrinter(hPrinter)
                PrintRaw = True
            Catch ex As Exception
                MsgBox("Error occurred: " & ex.ToString)
                PrintRaw = False
            Finally
                ' ----- Get rid of the special ANSI version.
                Marshal.FreeCoTaskMem(dataToSend)
            End Try
        End Function
    End Class
#End Region

#Region "MessagePanel"
    'Indicates current Message panel to add controls to 
    Private _CurrentMessagePanelName As String = Nothing

    'Used to give unique control names such as label1, label2 etc
    Private _MessagePanelsAddedCount As Integer = 0

    'Add Message panel to flow layout panel
    Public Sub CreateMessagePanel(Title As String, Message As String, Optional ByVal Tag As String = "")
        Dim MessagePanel As New Guna2Panel
        MessagePanel.Width = MdifrmMain.flpMessage.Width - 30
        MessagePanel.BorderRadius = 10
        MessagePanel.UseTransparentBackground = True
        MessagePanel.FillColor = Color.White
        'Add panel to flow layout panel
        MdifrmMain.flpMessage.Controls.Add(MessagePanel)

        'Update panel variables
        _CurrentMessagePanelName = MessagePanel.Name
        _MessagePanelsAddedCount += 1

        Dim MessageDeleteButton As New PictureBox

        'Set button properties
        With MessageDeleteButton
            .AutoSize = False
            .Size = New Size(16, 16)
            .Location = New Point(MessagePanel.Width - MessageDeleteButton.Width - 5, 5)
            .Image = My.Resources.close
            .Name = "btnMessageDelete" + (_MessagePanelsAddedCount).ToString
        End With

        'Add button to panel 
        For Each controlObject As Control In MdifrmMain.flpMessage.Controls
            If controlObject.Name = MessagePanel.Name Then
                controlObject.Controls.Add(MessageDeleteButton)
            End If
        Next

        'Add handler for click events
        AddHandler MessageDeleteButton.Click, AddressOf DynamicButton_Click

        Dim MessageTitleLabel As New Label

        With MessageTitleLabel
            .AutoSize = True
            .Font = New Font("Calibri", 15, FontStyle.Bold)
            .Location = New Point(5, 5)
            .Name = "lblTitle" + (_MessagePanelsAddedCount).ToString
            .Text = Title
            .ForeColor = Color.Black
            .Width = MessagePanel.Width - 10
            .MaximumSize = New Size(MessagePanel.Width - 10, 0)
        End With

        For Each controlObject As Control In MdifrmMain.flpMessage.Controls
            If controlObject.Name = MessagePanel.Name Then
                controlObject.Controls.Add(MessageTitleLabel)
            End If
        Next

        Dim MessageLabel As New Label

        With MessageLabel
            .AutoSize = True
            .Font = New Font("Calibri", 12, FontStyle.Regular)
            .Location = New Point(5, Int(MessageTitleLabel.Top) + Int(MessageTitleLabel.Height) + 5)
            .Name = "lblMessage" + (_MessagePanelsAddedCount).ToString
            .Text = Message
            .ForeColor = Color.DarkGray
            .Width = MessagePanel.Width - 10
            .MaximumSize = New Size(MessagePanel.Width - 10, 0)
        End With

        'Add button to panel 
        For Each controlObject As Control In MdifrmMain.flpMessage.Controls
            If controlObject.Name = MessagePanel.Name Then
                controlObject.Controls.Add(MessageLabel)
            End If
        Next

        'Set panel properties
        With MessagePanel
            .BackColor = Color.White
            .Height = Int(MessageLabel.Top) + Int(MessageLabel.Height) + 5
            .Name = "pnlMessage" + (_MessagePanelsAddedCount).ToString
            .Tag = Tag
        End With
    End Sub

    Public Sub CreateSMSMsgPanel(MsgNo As Integer)
        If MsgNo = Nothing Then Exit Sub
        Dim MessagePanel As New Guna2Panel
        MessagePanel.Width = MdifrmMain.flpMessage.Width - 30
        MessagePanel.BorderRadius = 10
        MessagePanel.UseTransparentBackground = True
        MessagePanel.FillColor = Color.White
        MessagePanel.Width = MdifrmMain.flpMessage.Width - 30
        'Add panel to flow layout panel
        MdifrmMain.flpMessage.Controls.Add(MessagePanel)

        'Update panel variables
        _CurrentMessagePanelName = MessagePanel.Name
        _MessagePanelsAddedCount += 1

        Dim MessageDeleteButton As New PictureBox

        'Set button properties
        With MessageDeleteButton
            .AutoSize = False
            .Size = New Size(16, 16)
            .Location = New Point(MessagePanel.Width - MessageDeleteButton.Width - 5, 5)
            .Image = My.Resources.close
            .Name = "btnMessageDelete" + (_MessagePanelsAddedCount).ToString
        End With

        'Add button to panel 
        For Each controlObject As Control In MdifrmMain.flpMessage.Controls
            If controlObject.Name = MessagePanel.Name Then
                controlObject.Controls.Add(MessageDeleteButton)
            End If
        Next

        'Add handler for click events
        AddHandler MessageDeleteButton.Click, AddressOf DynamicButton_Click

        Dim MessageTitleLabel As New Label

        With MessageTitleLabel
            .AutoSize = True
            .Font = New Font("Calibri", 15, FontStyle.Bold)
            .Location = New Point(5, 5)
            .Name = "lblTitle" + (_MessagePanelsAddedCount).ToString
            .Text = "Message එක Send කරන්නද?"
            .ForeColor = Color.Black
            .MaximumSize = New Size(MessagePanel.Width - 10, 0)
        End With

        For Each controlObject As Control In MdifrmMain.flpMessage.Controls
            If controlObject.Name = MessagePanel.Name Then
                controlObject.Controls.Add(MessageTitleLabel)
            End If
        Next

        Dim MessageLabel As New Label

        With MessageLabel
            .AutoSize = True
            .Font = New Font("Calibri", 12, FontStyle.Regular)
            .Location = New Point(5, Int(MessageTitleLabel.Top) + Int(MessageTitleLabel.Height) + 5)
            .Name = "lblMessage" + (_MessagePanelsAddedCount).ToString
            .Text = "Telephone No: " & GetStrfromRelatedfield("Select CuTelNo from Message Where MsgNo=" & MsgNo, "CuTelNo") &
                vbCrLf & "Message: " & GetStrfromRelatedfield("Select Message from Message Where MsgNo=" & MsgNo, "Message")
            .ForeColor = Color.DarkGray
            .MaximumSize = New Size(MessagePanel.Width - 10, 0)
        End With

        'Add button to panel 
        For Each controlObject As Control In MdifrmMain.flpMessage.Controls
            If controlObject.Name = MessagePanel.Name Then
                controlObject.Controls.Add(MessageLabel)
            End If
        Next

        Dim btnPrimary As New Button
        With btnPrimary
            .AutoSize = True
            .Font = New Font("Calibri", 12, FontStyle.Regular)
            .Location = New Point(5, Int(MessageLabel.Top) + Int(MessageLabel.Height) + 5)
            .Name = "btnPrimary" + (_MessagePanelsAddedCount).ToString
            .Text = "Send"
            .Tag = MsgNo
            .MaximumSize = New Size(((MessagePanel.Width - 10) / 2) - 5, 0)
        End With

        'Add button to panel 
        For Each controlObject As Control In MdifrmMain.flpMessage.Controls
            If controlObject.Name = MessagePanel.Name Then
                controlObject.Controls.Add(btnPrimary)
            End If
        Next
        'Add handler for click events
        AddHandler btnPrimary.Click, AddressOf btnPrimary_Click

        Dim btnSecondary As New Button
        With btnSecondary
            .AutoSize = True
            .Font = New Font("Calibri", 12, FontStyle.Regular)
            .Location = New Point(Int(btnPrimary.Left) + Int(btnPrimary.Width) + 5, Int(MessageLabel.Top) + Int(MessageLabel.Height) + 5)
            .Name = "btnSecondary" + (_MessagePanelsAddedCount).ToString
            .Text = "Don't Send"
            .Tag = MsgNo
            .MaximumSize = New Size(((MessagePanel.Width - 10) / 2) - 5, 0)
        End With

        'Add button to panel 
        For Each controlObject As Control In MdifrmMain.flpMessage.Controls
            If controlObject.Name = MessagePanel.Name Then
                controlObject.Controls.Add(btnSecondary)
            End If
        Next
        'Add handler for click events
        AddHandler btnSecondary.Click, AddressOf btnSecondary_Click

        'Set panel properties
        With MessagePanel
            .BackColor = Color.White
            .Height = Int(btnPrimary.Top) + Int(btnPrimary.Height) + 5
            .Name = "pnlMessage" + (_MessagePanelsAddedCount).ToString
            .Tag = "SendSMSConfirmation"
        End With
    End Sub

    Private Sub btnSecondary_Click(sender As Object, e As EventArgs)
        'Dim i As Integer = If(IsNumeric(sender.name.ToString.Replace("btnPrimary", "")), Int(sender.name.ToString.Replace("btnPrimary", "")), Nothing)
        CMDUPDATE("Update Message set Status='Declined' Where MsgNo=" & sender.tag)
        DynamicButton_Click(sender, e)
    End Sub

    Private Sub btnPrimary_Click(sender As Object, e As EventArgs)
        CMDUPDATE("Update Message set Status='Confirmed' Where MsgNo=" & sender.tag)
        DynamicButton_Click(sender, e)
    End Sub
    'Remove handlers and Message panel 
    Public Sub DynamicButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim parentPanelName As String

        parentPanelName = Nothing

        'Remove handler from sender
        For Each controlObj As Control In MdifrmMain.flpMessage.Controls
            For Each childControlObj As Control In controlObj.Controls
                If childControlObj.Name = sender.name Then
                    RemoveHandler childControlObj.Click, AddressOf DynamicButton_Click
                    parentPanelName = childControlObj.Parent.Name
                End If
            Next
            If controlObj.Name = parentPanelName Then
                MdifrmMain.flpMessage.Controls.Remove(controlObj)
                controlObj.Dispose()
            End If
        Next

    End Sub
#End Region

    Public Function CheckForInternetConnection() As Boolean
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("https://www.lasertec.lk")
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function

    <Obsolete>
    Public Function GetMD5Encoder(theInput As String) As String
        ' Convert to byte array and get hash
        Dim dbytes As Byte() = Encoding.UTF8.GetBytes(FormsAuthentication.HashPasswordForStoringInConfigFile(theInput, "MD5"))

        Return Convert.ToBase64String(dbytes)
    End Function

#Region "Encode&Decode"
    Public NotInheritable Class Simple3Des
        Private TripleDes As New TripleDESCryptoServiceProvider

        Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()

            Dim sha1 As New SHA1CryptoServiceProvider

            ' Hash the key.
            Dim keyBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(key)
            Dim hash() As Byte = sha1.ComputeHash(keyBytes)

            ' Truncate or pad the hash.
            ReDim Preserve hash(length - 1)
            Return hash
        End Function

        Sub New(ByVal key As String)
            ' Initialize the crypto provider.
            TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
            TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
        End Sub

        Private Function EncryptData(ByVal plaintext As String) As String

            ' Convert the plaintext string to a byte array. 
            Dim plaintextBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(plaintext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the encoder to write to the stream. 
            Dim encStream As New CryptoStream(ms,
                TripleDes.CreateEncryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
            encStream.FlushFinalBlock()

            ' Convert the encrypted stream to a printable string. 


            Return Convert.ToBase64String(ms.ToArray)

        End Function

        Private Function DecryptData(ByVal encryptedtext As String) As String

            ' Convert the encrypted text string to a byte array. 
            Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the decoder to write to the stream. 
            Dim decStream As New CryptoStream(ms,
                TripleDes.CreateDecryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()

            ' Convert the plaintext stream to a string. 
            Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
        End Function

        Public Function Decode(cipher As String) As String
            Try
                If cipher <> "" Then
                    Return DecryptData(cipher)
                Else
                    Return ""
                End If
            Catch ex As CryptographicException
                Throw New Exception(ex.Message)
            End Try
        End Function

        Public Function Encode(txt As String) As String
            Try
                Return EncryptData(txt)
            Catch ex As CryptographicException
                Throw New Exception(ex.Message)
            End Try
        End Function
    End Class
#End Region

    Public Function mnustrpMENU() As ToolStripMenuItem
        Dim cmdStock As ToolStripMenuItem = New ToolStripMenuItem("Stock")
        Dim cmdCustomer As ToolStripMenuItem = New ToolStripMenuItem("Customer")
        Dim cmdSupplier As ToolStripMenuItem = New ToolStripMenuItem("Supplier")
        Dim cmdTechnician As ToolStripMenuItem = New ToolStripMenuItem("Technician")
        Dim cmdProduct As ToolStripMenuItem = New ToolStripMenuItem("Product")
        Dim cmdSale As ToolStripMenuItem = New ToolStripMenuItem("Sale")
        Dim cmdOpenSale As ToolStripMenuItem = New ToolStripMenuItem("Open Sale")
        Dim cmdNewSale As ToolStripMenuItem = New ToolStripMenuItem("New Sale")
        Dim cmdSupply As ToolStripMenuItem = New ToolStripMenuItem("Supply")
        Dim cmdReceive As ToolStripMenuItem = New ToolStripMenuItem("Receive")
        Dim cmdRepair As ToolStripMenuItem = New ToolStripMenuItem("Repair")
        Dim cmdDeliver As ToolStripMenuItem = New ToolStripMenuItem("Deliver")
        Dim cmdSalesRepair As ToolStripMenuItem = New ToolStripMenuItem("Sales Repair")
        Dim cmdTechnicianCost As ToolStripMenuItem = New ToolStripMenuItem("Technician Cost")
        Dim cmdTechnicianLoan As ToolStripMenuItem = New ToolStripMenuItem("Technician Loan")
        Dim cmdTechnicianSalary As ToolStripMenuItem = New ToolStripMenuItem("Technician Salary")
        Dim cmdCustomerLoan As ToolStripMenuItem = New ToolStripMenuItem("Customer Loan")
        Dim cmdSettlement As ToolStripMenuItem = New ToolStripMenuItem("Settlement")
        Dim cmdBarCodeGenerator As ToolStripMenuItem = New ToolStripMenuItem("Barcode Generator")
        Dim cmdTool As ToolStripMenuItem = New ToolStripMenuItem("Tools")
        Dim cmdOpenCashDrawer As ToolStripMenuItem = New ToolStripMenuItem("Open Cash Drawer")
        Dim sep1, sep2, sep3, sep4 As New ToolStripSeparator
        cmdStock.ShortcutKeys = Keys.F1
        AddHandler cmdStock.Click, AddressOf MdifrmMain.CmdStock_Click
        AddHandler cmdCustomer.Click, AddressOf MdifrmMain.CmdCustomer_Click
        AddHandler cmdSupplier.Click, AddressOf MdifrmMain.CmdSupplier_Click
        AddHandler cmdTechnician.Click, AddressOf MdifrmMain.cmdTechnician_Click
        AddHandler cmdProduct.Click, AddressOf MdifrmMain.CmdProduct_Click
        AddHandler cmdOpenSale.Click, AddressOf MdifrmMain.OpenSaleToolStripMenuItem_Click
        cmdOpenSale.ShortcutKeys = Keys.F2 + Keys.Control
        AddHandler cmdNewSale.Click, AddressOf MdifrmMain.NewSaleToolStripMenuItem_Click
        cmdNewSale.ShortcutKeys = Keys.F2
        AddHandler cmdSupply.Click, AddressOf MdifrmMain.CmdSupply_Click
        cmdReceive.ShortcutKeys = Keys.F3
        AddHandler cmdReceive.Click, AddressOf MdifrmMain.CmdReceive_Click
        cmdRepair.ShortcutKeys = Keys.F4
        AddHandler cmdRepair.Click, AddressOf MdifrmMain.cmdRepair_Click
        cmdDeliver.ShortcutKeys = Keys.F5
        AddHandler cmdDeliver.Click, AddressOf MdifrmMain.cmdDeliver_Click
        AddHandler cmdSalesRepair.Click, AddressOf MdifrmMain.cmdSalesRepair_Click
        cmdTechnicianCost.ShortcutKeys = Keys.F6
        AddHandler cmdTechnicianCost.Click, AddressOf MdifrmMain.cmdTechnicianCost_Click
        cmdTechnicianLoan.ShortcutKeys = Keys.F7
        AddHandler cmdTechnicianLoan.Click, AddressOf MdifrmMain.cmdTechnicianLoan_Click
        AddHandler cmdTechnicianSalary.Click, AddressOf MdifrmMain.cmdTechnicianSalary_Click
        AddHandler cmdCustomerLoan.Click, AddressOf MdifrmMain.cmdCustomerLoan_Click
        AddHandler cmdSettlement.Click, AddressOf MdifrmMain.CmdSettlement_Click
        AddHandler cmdBarCodeGenerator.Click, AddressOf MdifrmMain.BarCodeGeneratorToolStripMenuItem_Click
        cmdOpenCashDrawer.ShortcutKeys = Keys.F8
        AddHandler cmdOpenCashDrawer.Click, AddressOf OpenCashdrawer

        mnustrpMENU = New ToolStripMenuItem("MENU")
        mnustrpMENU.DropDownItems.Add(cmdStock)
        mnustrpMENU.DropDownItems.Add(cmdCustomer)
        mnustrpMENU.DropDownItems.Add(cmdSupplier)
        mnustrpMENU.DropDownItems.Add(cmdTechnician)
        mnustrpMENU.DropDownItems.Add(cmdProduct)
        mnustrpMENU.DropDownItems.Add(sep1)

        mnustrpMENU.DropDownItems.Add(cmdSale)
        cmdSale.DropDownItems.Add(cmdNewSale)
        cmdSale.DropDownItems.Add(cmdOpenSale)
        mnustrpMENU.DropDownItems.Add(cmdSupply)
        mnustrpMENU.DropDownItems.Add(sep2)

        mnustrpMENU.DropDownItems.Add(cmdReceive)
        mnustrpMENU.DropDownItems.Add(cmdRepair)
        mnustrpMENU.DropDownItems.Add(cmdDeliver)
        mnustrpMENU.DropDownItems.Add(cmdSalesRepair)
        mnustrpMENU.DropDownItems.Add(sep3)

        mnustrpMENU.DropDownItems.Add(cmdTechnicianCost)
        mnustrpMENU.DropDownItems.Add(cmdTechnicianLoan)
        mnustrpMENU.DropDownItems.Add(cmdTechnicianSalary)
        mnustrpMENU.DropDownItems.Add(cmdCustomerLoan)
        mnustrpMENU.DropDownItems.Add(cmdSettlement)
        mnustrpMENU.DropDownItems.Add(cmdBarCodeGenerator)
        mnustrpMENU.DropDownItems.Add(sep4)

        mnustrpMENU.DropDownItems.Add(cmdTool)
        cmdTool.DropDownItems.Add(cmdOpenCashDrawer)

        Return mnustrpMENU
    End Function

End Module
