﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmSettings
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmSettings))
        Me.tcSettings = New System.Windows.Forms.TabControl()
        Me.tpGeneral = New System.Windows.Forms.TabPage()
        Me.Guna2GroupBox4 = New Guna.UI2.WinForms.Guna2GroupBox()
        Me.txtBSBaudRate = New System.Windows.Forms.TextBox()
        Me.cmbBSCOMPort = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.chkBSCOMMode = New System.Windows.Forms.CheckBox()
        Me.Guna2GroupBox3 = New Guna.UI2.WinForms.Guna2GroupBox()
        Me.txtMAdminEmail = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.chkDeliveredEmailtoT = New System.Windows.Forms.CheckBox()
        Me.chkMSetEmail = New System.Windows.Forms.CheckBox()
        Me.tpDatabase = New System.Windows.Forms.TabPage()
        Me.txtDBPassword = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtDBLoc = New System.Windows.Forms.TextBox()
        Me.cmdDBLocation = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tpPrinter = New System.Windows.Forms.TabPage()
        Me.Guna2GroupBox2 = New Guna.UI2.WinForms.Guna2GroupBox()
        Me.txtBillPaperName = New System.Windows.Forms.TextBox()
        Me.txtBillPrinterName = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Guna2GroupBox1 = New Guna.UI2.WinForms.Guna2GroupBox()
        Me.txtStickerStockPaperName = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtStickerRepairPaperName = New System.Windows.Forms.TextBox()
        Me.txtStickerPrinterName = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.tpUserAccount = New System.Windows.Forms.TabPage()
        Me.grdUAUser = New System.Windows.Forms.DataGridView()
        Me.UAUNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.UAUserName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.UAType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.UAEmail = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.grpUAUser = New System.Windows.Forms.GroupBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cmbUAType = New System.Windows.Forms.ComboBox()
        Me.txtUAUNo = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtUACurrentPW = New System.Windows.Forms.TextBox()
        Me.lblUACurrentPW = New System.Windows.Forms.Label()
        Me.txtUAEmail = New System.Windows.Forms.TextBox()
        Me.lblUAEmail = New System.Windows.Forms.Label()
        Me.txtUAReenterPW = New System.Windows.Forms.TextBox()
        Me.lblUAReenterPW = New System.Windows.Forms.Label()
        Me.cmdUADelete = New System.Windows.Forms.Button()
        Me.cmdUASave = New System.Windows.Forms.Button()
        Me.cmdUANew = New System.Windows.Forms.Button()
        Me.txtUANewPW = New System.Windows.Forms.TextBox()
        Me.lblUANewPW = New System.Windows.Forms.Label()
        Me.txtUAUserName = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.grpUAAdmin = New System.Windows.Forms.GroupBox()
        Me.cmdUACheck = New System.Windows.Forms.Button()
        Me.txtUAAPassword = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtUAAUserName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdApply = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.ofdDatabase = New System.Windows.Forms.OpenFileDialog()
        Me.cmbDBProvider = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tcSettings.SuspendLayout()
        Me.tpGeneral.SuspendLayout()
        Me.Guna2GroupBox4.SuspendLayout()
        Me.Guna2GroupBox3.SuspendLayout()
        Me.tpDatabase.SuspendLayout()
        Me.tpPrinter.SuspendLayout()
        Me.Guna2GroupBox2.SuspendLayout()
        Me.Guna2GroupBox1.SuspendLayout()
        Me.tpUserAccount.SuspendLayout()
        CType(Me.grdUAUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpUAUser.SuspendLayout()
        Me.grpUAAdmin.SuspendLayout()
        Me.SuspendLayout()
        '
        'tcSettings
        '
        Me.tcSettings.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tcSettings.Controls.Add(Me.tpGeneral)
        Me.tcSettings.Controls.Add(Me.tpDatabase)
        Me.tcSettings.Controls.Add(Me.tpPrinter)
        Me.tcSettings.Controls.Add(Me.tpUserAccount)
        Me.tcSettings.HotTrack = True
        Me.tcSettings.Location = New System.Drawing.Point(12, 12)
        Me.tcSettings.Name = "tcSettings"
        Me.tcSettings.SelectedIndex = 0
        Me.tcSettings.Size = New System.Drawing.Size(713, 379)
        Me.tcSettings.TabIndex = 0
        '
        'tpGeneral
        '
        Me.tpGeneral.Controls.Add(Me.Guna2GroupBox4)
        Me.tpGeneral.Controls.Add(Me.Guna2GroupBox3)
        Me.tpGeneral.Location = New System.Drawing.Point(4, 22)
        Me.tpGeneral.Name = "tpGeneral"
        Me.tpGeneral.Padding = New System.Windows.Forms.Padding(3)
        Me.tpGeneral.Size = New System.Drawing.Size(705, 353)
        Me.tpGeneral.TabIndex = 1
        Me.tpGeneral.Text = "General"
        Me.tpGeneral.UseVisualStyleBackColor = True
        '
        'Guna2GroupBox4
        '
        Me.Guna2GroupBox4.Controls.Add(Me.txtBSBaudRate)
        Me.Guna2GroupBox4.Controls.Add(Me.cmbBSCOMPort)
        Me.Guna2GroupBox4.Controls.Add(Me.Label6)
        Me.Guna2GroupBox4.Controls.Add(Me.Label2)
        Me.Guna2GroupBox4.Controls.Add(Me.chkBSCOMMode)
        Me.Guna2GroupBox4.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2GroupBox4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(125, Byte), Integer), CType(CType(137, Byte), Integer), CType(CType(149, Byte), Integer))
        Me.Guna2GroupBox4.Location = New System.Drawing.Point(6, 196)
        Me.Guna2GroupBox4.Name = "Guna2GroupBox4"
        Me.Guna2GroupBox4.ShadowDecoration.Parent = Me.Guna2GroupBox4
        Me.Guna2GroupBox4.Size = New System.Drawing.Size(365, 108)
        Me.Guna2GroupBox4.TabIndex = 22
        Me.Guna2GroupBox4.Text = "Barcode Scanner Info"
        '
        'txtBSBaudRate
        '
        Me.txtBSBaudRate.Location = New System.Drawing.Point(243, 74)
        Me.txtBSBaudRate.Name = "txtBSBaudRate"
        Me.txtBSBaudRate.Size = New System.Drawing.Size(81, 23)
        Me.txtBSBaudRate.TabIndex = 4
        '
        'cmbBSCOMPort
        '
        Me.cmbBSCOMPort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbBSCOMPort.FormattingEnabled = True
        Me.cmbBSCOMPort.Location = New System.Drawing.Point(60, 74)
        Me.cmbBSCOMPort.Name = "cmbBSCOMPort"
        Me.cmbBSCOMPort.Size = New System.Drawing.Size(106, 23)
        Me.cmbBSCOMPort.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(174, 77)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(63, 15)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Baud Rate:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 77)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 15)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Port :"
        '
        'chkBSCOMMode
        '
        Me.chkBSCOMMode.AutoSize = True
        Me.chkBSCOMMode.Location = New System.Drawing.Point(11, 45)
        Me.chkBSCOMMode.Name = "chkBSCOMMode"
        Me.chkBSCOMMode.Size = New System.Drawing.Size(88, 19)
        Me.chkBSCOMMode.TabIndex = 0
        Me.chkBSCOMMode.Text = "COM Mode"
        Me.chkBSCOMMode.UseVisualStyleBackColor = True
        '
        'Guna2GroupBox3
        '
        Me.Guna2GroupBox3.Controls.Add(Me.txtMAdminEmail)
        Me.Guna2GroupBox3.Controls.Add(Me.Label15)
        Me.Guna2GroupBox3.Controls.Add(Me.chkDeliveredEmailtoT)
        Me.Guna2GroupBox3.Controls.Add(Me.chkMSetEmail)
        Me.Guna2GroupBox3.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2GroupBox3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(125, Byte), Integer), CType(CType(137, Byte), Integer), CType(CType(149, Byte), Integer))
        Me.Guna2GroupBox3.Location = New System.Drawing.Point(6, 6)
        Me.Guna2GroupBox3.Name = "Guna2GroupBox3"
        Me.Guna2GroupBox3.ShadowDecoration.Parent = Me.Guna2GroupBox3
        Me.Guna2GroupBox3.Size = New System.Drawing.Size(365, 157)
        Me.Guna2GroupBox3.TabIndex = 21
        Me.Guna2GroupBox3.Text = "Email Info"
        '
        'txtMAdminEmail
        '
        Me.txtMAdminEmail.Location = New System.Drawing.Point(94, 48)
        Me.txtMAdminEmail.Name = "txtMAdminEmail"
        Me.txtMAdminEmail.Size = New System.Drawing.Size(263, 23)
        Me.txtMAdminEmail.TabIndex = 36
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(10, 51)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(78, 15)
        Me.Label15.TabIndex = 35
        Me.Label15.Text = "Admin Email:"
        '
        'chkDeliveredEmailtoT
        '
        Me.chkDeliveredEmailtoT.Location = New System.Drawing.Point(7, 115)
        Me.chkDeliveredEmailtoT.Name = "chkDeliveredEmailtoT"
        Me.chkDeliveredEmailtoT.Size = New System.Drawing.Size(340, 34)
        Me.chkDeliveredEmailtoT.TabIndex = 34
        Me.chkDeliveredEmailtoT.Text = "Automatically send Delivered Result to Technician"
        Me.chkDeliveredEmailtoT.UseVisualStyleBackColor = True
        '
        'chkMSetEmail
        '
        Me.chkMSetEmail.Location = New System.Drawing.Point(7, 75)
        Me.chkMSetEmail.Name = "chkMSetEmail"
        Me.chkMSetEmail.Size = New System.Drawing.Size(350, 34)
        Me.chkMSetEmail.TabIndex = 30
        Me.chkMSetEmail.Text = "Send Settlement, Technician Cost,Technician Loan and Transaction to Admin via Ema" &
    "il" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.chkMSetEmail.UseVisualStyleBackColor = True
        '
        'tpDatabase
        '
        Me.tpDatabase.Controls.Add(Me.Label10)
        Me.tpDatabase.Controls.Add(Me.cmbDBProvider)
        Me.tpDatabase.Controls.Add(Me.txtDBPassword)
        Me.tpDatabase.Controls.Add(Me.Label7)
        Me.tpDatabase.Controls.Add(Me.txtDBLoc)
        Me.tpDatabase.Controls.Add(Me.cmdDBLocation)
        Me.tpDatabase.Controls.Add(Me.Label1)
        Me.tpDatabase.Location = New System.Drawing.Point(4, 22)
        Me.tpDatabase.Name = "tpDatabase"
        Me.tpDatabase.Padding = New System.Windows.Forms.Padding(3)
        Me.tpDatabase.Size = New System.Drawing.Size(705, 353)
        Me.tpDatabase.TabIndex = 0
        Me.tpDatabase.Text = "Database"
        Me.tpDatabase.UseVisualStyleBackColor = True
        '
        'txtDBPassword
        '
        Me.txtDBPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDBPassword.Location = New System.Drawing.Point(133, 33)
        Me.txtDBPassword.Name = "txtDBPassword"
        Me.txtDBPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtDBPassword.Size = New System.Drawing.Size(150, 21)
        Me.txtDBPassword.TabIndex = 8
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 36)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(121, 13)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Set Database Password:"
        '
        'txtDBLoc
        '
        Me.txtDBLoc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDBLoc.Location = New System.Drawing.Point(110, 7)
        Me.txtDBLoc.Name = "txtDBLoc"
        Me.txtDBLoc.Size = New System.Drawing.Size(558, 21)
        Me.txtDBLoc.TabIndex = 6
        '
        'cmdDBLocation
        '
        Me.cmdDBLocation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDBLocation.Location = New System.Drawing.Point(674, 7)
        Me.cmdDBLocation.Name = "cmdDBLocation"
        Me.cmdDBLocation.Size = New System.Drawing.Size(25, 21)
        Me.cmdDBLocation.TabIndex = 2
        Me.cmdDBLocation.Text = "..."
        Me.cmdDBLocation.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Database Location:"
        '
        'tpPrinter
        '
        Me.tpPrinter.Controls.Add(Me.Guna2GroupBox2)
        Me.tpPrinter.Controls.Add(Me.Guna2GroupBox1)
        Me.tpPrinter.Location = New System.Drawing.Point(4, 22)
        Me.tpPrinter.Name = "tpPrinter"
        Me.tpPrinter.Padding = New System.Windows.Forms.Padding(3)
        Me.tpPrinter.Size = New System.Drawing.Size(705, 353)
        Me.tpPrinter.TabIndex = 3
        Me.tpPrinter.Text = "Print"
        Me.tpPrinter.UseVisualStyleBackColor = True
        '
        'Guna2GroupBox2
        '
        Me.Guna2GroupBox2.Controls.Add(Me.txtBillPaperName)
        Me.Guna2GroupBox2.Controls.Add(Me.txtBillPrinterName)
        Me.Guna2GroupBox2.Controls.Add(Me.Label22)
        Me.Guna2GroupBox2.Controls.Add(Me.Label23)
        Me.Guna2GroupBox2.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2GroupBox2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(125, Byte), Integer), CType(CType(137, Byte), Integer), CType(CType(149, Byte), Integer))
        Me.Guna2GroupBox2.Location = New System.Drawing.Point(6, 152)
        Me.Guna2GroupBox2.Name = "Guna2GroupBox2"
        Me.Guna2GroupBox2.ShadowDecoration.Parent = Me.Guna2GroupBox2
        Me.Guna2GroupBox2.Size = New System.Drawing.Size(308, 104)
        Me.Guna2GroupBox2.TabIndex = 1
        Me.Guna2GroupBox2.Text = "Bill Printer Info"
        '
        'txtBillPaperName
        '
        Me.txtBillPaperName.Location = New System.Drawing.Point(94, 72)
        Me.txtBillPaperName.Name = "txtBillPaperName"
        Me.txtBillPaperName.Size = New System.Drawing.Size(200, 23)
        Me.txtBillPaperName.TabIndex = 35
        '
        'txtBillPrinterName
        '
        Me.txtBillPrinterName.Location = New System.Drawing.Point(94, 45)
        Me.txtBillPrinterName.Name = "txtBillPrinterName"
        Me.txtBillPrinterName.Size = New System.Drawing.Size(200, 23)
        Me.txtBillPrinterName.TabIndex = 34
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(3, 75)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(75, 15)
        Me.Label22.TabIndex = 33
        Me.Label22.Text = "Paper Name:"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(3, 48)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(80, 15)
        Me.Label23.TabIndex = 32
        Me.Label23.Text = "Printer Name:"
        '
        'Guna2GroupBox1
        '
        Me.Guna2GroupBox1.Controls.Add(Me.txtStickerStockPaperName)
        Me.Guna2GroupBox1.Controls.Add(Me.Label24)
        Me.Guna2GroupBox1.Controls.Add(Me.txtStickerRepairPaperName)
        Me.Guna2GroupBox1.Controls.Add(Me.txtStickerPrinterName)
        Me.Guna2GroupBox1.Controls.Add(Me.Label18)
        Me.Guna2GroupBox1.Controls.Add(Me.Label21)
        Me.Guna2GroupBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.Guna2GroupBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(125, Byte), Integer), CType(CType(137, Byte), Integer), CType(CType(149, Byte), Integer))
        Me.Guna2GroupBox1.Location = New System.Drawing.Point(6, 6)
        Me.Guna2GroupBox1.Name = "Guna2GroupBox1"
        Me.Guna2GroupBox1.ShadowDecoration.Parent = Me.Guna2GroupBox1
        Me.Guna2GroupBox1.Size = New System.Drawing.Size(308, 140)
        Me.Guna2GroupBox1.TabIndex = 0
        Me.Guna2GroupBox1.Text = "Sticker Printer Info"
        '
        'txtStickerStockPaperName
        '
        Me.txtStickerStockPaperName.Location = New System.Drawing.Point(154, 101)
        Me.txtStickerStockPaperName.Name = "txtStickerStockPaperName"
        Me.txtStickerStockPaperName.Size = New System.Drawing.Size(140, 23)
        Me.txtStickerStockPaperName.TabIndex = 37
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(3, 104)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(145, 15)
        Me.Label24.TabIndex = 36
        Me.Label24.Text = "Stock Sticker Paper Name:"
        '
        'txtStickerRepairPaperName
        '
        Me.txtStickerRepairPaperName.Location = New System.Drawing.Point(158, 72)
        Me.txtStickerRepairPaperName.Name = "txtStickerRepairPaperName"
        Me.txtStickerRepairPaperName.Size = New System.Drawing.Size(136, 23)
        Me.txtStickerRepairPaperName.TabIndex = 35
        '
        'txtStickerPrinterName
        '
        Me.txtStickerPrinterName.Location = New System.Drawing.Point(94, 45)
        Me.txtStickerPrinterName.Name = "txtStickerPrinterName"
        Me.txtStickerPrinterName.Size = New System.Drawing.Size(200, 23)
        Me.txtStickerPrinterName.TabIndex = 34
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(3, 75)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(149, 15)
        Me.Label18.TabIndex = 33
        Me.Label18.Text = "Repair Sticker Paper Name:"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(3, 48)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(80, 15)
        Me.Label21.TabIndex = 32
        Me.Label21.Text = "Printer Name:"
        '
        'tpUserAccount
        '
        Me.tpUserAccount.Controls.Add(Me.grdUAUser)
        Me.tpUserAccount.Controls.Add(Me.grpUAUser)
        Me.tpUserAccount.Controls.Add(Me.grpUAAdmin)
        Me.tpUserAccount.Location = New System.Drawing.Point(4, 22)
        Me.tpUserAccount.Name = "tpUserAccount"
        Me.tpUserAccount.Size = New System.Drawing.Size(705, 353)
        Me.tpUserAccount.TabIndex = 2
        Me.tpUserAccount.Text = "User Account"
        Me.tpUserAccount.UseVisualStyleBackColor = True
        '
        'grdUAUser
        '
        Me.grdUAUser.AllowUserToAddRows = False
        Me.grdUAUser.AllowUserToDeleteRows = False
        Me.grdUAUser.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdUAUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdUAUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.UAUNo, Me.UAUserName, Me.UAType, Me.UAEmail})
        Me.grdUAUser.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.grdUAUser.Enabled = False
        Me.grdUAUser.Location = New System.Drawing.Point(295, 6)
        Me.grdUAUser.Name = "grdUAUser"
        Me.grdUAUser.ReadOnly = True
        Me.grdUAUser.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdUAUser.Size = New System.Drawing.Size(404, 340)
        Me.grdUAUser.TabIndex = 2
        '
        'UAUNo
        '
        Me.UAUNo.DataPropertyName = "UNo"
        Me.UAUNo.HeaderText = "User No"
        Me.UAUNo.Name = "UAUNo"
        Me.UAUNo.ReadOnly = True
        '
        'UAUserName
        '
        Me.UAUserName.DataPropertyName = "UserName"
        Me.UAUserName.HeaderText = "User Name"
        Me.UAUserName.Name = "UAUserName"
        Me.UAUserName.ReadOnly = True
        Me.UAUserName.Width = 61
        '
        'UAType
        '
        Me.UAType.DataPropertyName = "Type"
        Me.UAType.HeaderText = "User Type"
        Me.UAType.Name = "UAType"
        Me.UAType.ReadOnly = True
        '
        'UAEmail
        '
        Me.UAEmail.DataPropertyName = "Email"
        Me.UAEmail.HeaderText = "User Email"
        Me.UAEmail.Name = "UAEmail"
        Me.UAEmail.ReadOnly = True
        '
        'grpUAUser
        '
        Me.grpUAUser.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpUAUser.Controls.Add(Me.Label9)
        Me.grpUAUser.Controls.Add(Me.cmbUAType)
        Me.grpUAUser.Controls.Add(Me.txtUAUNo)
        Me.grpUAUser.Controls.Add(Me.Label5)
        Me.grpUAUser.Controls.Add(Me.txtUACurrentPW)
        Me.grpUAUser.Controls.Add(Me.lblUACurrentPW)
        Me.grpUAUser.Controls.Add(Me.txtUAEmail)
        Me.grpUAUser.Controls.Add(Me.lblUAEmail)
        Me.grpUAUser.Controls.Add(Me.txtUAReenterPW)
        Me.grpUAUser.Controls.Add(Me.lblUAReenterPW)
        Me.grpUAUser.Controls.Add(Me.cmdUADelete)
        Me.grpUAUser.Controls.Add(Me.cmdUASave)
        Me.grpUAUser.Controls.Add(Me.cmdUANew)
        Me.grpUAUser.Controls.Add(Me.txtUANewPW)
        Me.grpUAUser.Controls.Add(Me.lblUANewPW)
        Me.grpUAUser.Controls.Add(Me.txtUAUserName)
        Me.grpUAUser.Controls.Add(Me.Label8)
        Me.grpUAUser.Enabled = False
        Me.grpUAUser.Location = New System.Drawing.Point(8, 119)
        Me.grpUAUser.Name = "grpUAUser"
        Me.grpUAUser.Size = New System.Drawing.Size(281, 228)
        Me.grpUAUser.TabIndex = 1
        Me.grpUAUser.TabStop = False
        Me.grpUAUser.Text = "User Details"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(105, 23)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(53, 13)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "User Type"
        '
        'cmbUAType
        '
        Me.cmbUAType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbUAType.FormattingEnabled = True
        Me.cmbUAType.Items.AddRange(New Object() {"Cashier", "Admin"})
        Me.cmbUAType.Location = New System.Drawing.Point(164, 20)
        Me.cmbUAType.Name = "cmbUAType"
        Me.cmbUAType.Size = New System.Drawing.Size(107, 21)
        Me.cmbUAType.TabIndex = 18
        '
        'txtUAUNo
        '
        Me.txtUAUNo.Enabled = False
        Me.txtUAUNo.Location = New System.Drawing.Point(55, 20)
        Me.txtUAUNo.Name = "txtUAUNo"
        Me.txtUAUNo.Size = New System.Drawing.Size(44, 21)
        Me.txtUAUNo.TabIndex = 17
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 23)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(44, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "User No"
        '
        'txtUACurrentPW
        '
        Me.txtUACurrentPW.Location = New System.Drawing.Point(105, 74)
        Me.txtUACurrentPW.Name = "txtUACurrentPW"
        Me.txtUACurrentPW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtUACurrentPW.Size = New System.Drawing.Size(166, 21)
        Me.txtUACurrentPW.TabIndex = 15
        '
        'lblUACurrentPW
        '
        Me.lblUACurrentPW.AutoSize = True
        Me.lblUACurrentPW.Location = New System.Drawing.Point(6, 77)
        Me.lblUACurrentPW.Name = "lblUACurrentPW"
        Me.lblUACurrentPW.Size = New System.Drawing.Size(91, 13)
        Me.lblUACurrentPW.TabIndex = 14
        Me.lblUACurrentPW.Text = "Current Password"
        '
        'txtUAEmail
        '
        Me.txtUAEmail.Location = New System.Drawing.Point(44, 155)
        Me.txtUAEmail.Name = "txtUAEmail"
        Me.txtUAEmail.Size = New System.Drawing.Size(227, 21)
        Me.txtUAEmail.TabIndex = 13
        '
        'lblUAEmail
        '
        Me.lblUAEmail.AutoSize = True
        Me.lblUAEmail.Location = New System.Drawing.Point(5, 158)
        Me.lblUAEmail.Name = "lblUAEmail"
        Me.lblUAEmail.Size = New System.Drawing.Size(33, 13)
        Me.lblUAEmail.TabIndex = 12
        Me.lblUAEmail.Text = "Email"
        '
        'txtUAReenterPW
        '
        Me.txtUAReenterPW.Location = New System.Drawing.Point(105, 128)
        Me.txtUAReenterPW.Name = "txtUAReenterPW"
        Me.txtUAReenterPW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtUAReenterPW.Size = New System.Drawing.Size(166, 21)
        Me.txtUAReenterPW.TabIndex = 11
        '
        'lblUAReenterPW
        '
        Me.lblUAReenterPW.AutoSize = True
        Me.lblUAReenterPW.Location = New System.Drawing.Point(6, 131)
        Me.lblUAReenterPW.Name = "lblUAReenterPW"
        Me.lblUAReenterPW.Size = New System.Drawing.Size(93, 13)
        Me.lblUAReenterPW.TabIndex = 10
        Me.lblUAReenterPW.Text = "Reenter Password"
        '
        'cmdUADelete
        '
        Me.cmdUADelete.Image = Global.LASERSystem.My.Resources.Resources.Delete
        Me.cmdUADelete.Location = New System.Drawing.Point(208, 182)
        Me.cmdUADelete.Name = "cmdUADelete"
        Me.cmdUADelete.Size = New System.Drawing.Size(67, 27)
        Me.cmdUADelete.TabIndex = 9
        Me.cmdUADelete.Text = "Delete"
        Me.cmdUADelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdUADelete.UseVisualStyleBackColor = True
        '
        'cmdUASave
        '
        Me.cmdUASave.Image = Global.LASERSystem.My.Resources.Resources.Save
        Me.cmdUASave.Location = New System.Drawing.Point(135, 182)
        Me.cmdUASave.Name = "cmdUASave"
        Me.cmdUASave.Size = New System.Drawing.Size(67, 27)
        Me.cmdUASave.TabIndex = 8
        Me.cmdUASave.Text = "Save"
        Me.cmdUASave.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdUASave.UseVisualStyleBackColor = True
        '
        'cmdUANew
        '
        Me.cmdUANew.Image = Global.LASERSystem.My.Resources.Resources._new
        Me.cmdUANew.Location = New System.Drawing.Point(62, 182)
        Me.cmdUANew.Name = "cmdUANew"
        Me.cmdUANew.Size = New System.Drawing.Size(67, 27)
        Me.cmdUANew.TabIndex = 7
        Me.cmdUANew.Text = "New"
        Me.cmdUANew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdUANew.UseVisualStyleBackColor = True
        '
        'txtUANewPW
        '
        Me.txtUANewPW.Location = New System.Drawing.Point(105, 101)
        Me.txtUANewPW.Name = "txtUANewPW"
        Me.txtUANewPW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtUANewPW.Size = New System.Drawing.Size(166, 21)
        Me.txtUANewPW.TabIndex = 3
        '
        'lblUANewPW
        '
        Me.lblUANewPW.AutoSize = True
        Me.lblUANewPW.Location = New System.Drawing.Point(6, 104)
        Me.lblUANewPW.Name = "lblUANewPW"
        Me.lblUANewPW.Size = New System.Drawing.Size(76, 13)
        Me.lblUANewPW.TabIndex = 2
        Me.lblUANewPW.Text = "New Password"
        '
        'txtUAUserName
        '
        Me.txtUAUserName.Location = New System.Drawing.Point(70, 47)
        Me.txtUAUserName.Name = "txtUAUserName"
        Me.txtUAUserName.Size = New System.Drawing.Size(201, 21)
        Me.txtUAUserName.TabIndex = 1
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(5, 50)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(59, 13)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "User Name"
        '
        'grpUAAdmin
        '
        Me.grpUAAdmin.Controls.Add(Me.cmdUACheck)
        Me.grpUAAdmin.Controls.Add(Me.txtUAAPassword)
        Me.grpUAAdmin.Controls.Add(Me.Label4)
        Me.grpUAAdmin.Controls.Add(Me.txtUAAUserName)
        Me.grpUAAdmin.Controls.Add(Me.Label3)
        Me.grpUAAdmin.Location = New System.Drawing.Point(8, 6)
        Me.grpUAAdmin.Name = "grpUAAdmin"
        Me.grpUAAdmin.Size = New System.Drawing.Size(281, 107)
        Me.grpUAAdmin.TabIndex = 0
        Me.grpUAAdmin.TabStop = False
        Me.grpUAAdmin.Text = "Admin Details"
        '
        'cmdUACheck
        '
        Me.cmdUACheck.Location = New System.Drawing.Point(204, 73)
        Me.cmdUACheck.Name = "cmdUACheck"
        Me.cmdUACheck.Size = New System.Drawing.Size(67, 27)
        Me.cmdUACheck.TabIndex = 7
        Me.cmdUACheck.Text = "Check"
        Me.cmdUACheck.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdUACheck.UseVisualStyleBackColor = True
        '
        'txtUAAPassword
        '
        Me.txtUAAPassword.Location = New System.Drawing.Point(76, 47)
        Me.txtUAAPassword.Name = "txtUAAPassword"
        Me.txtUAAPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtUAAPassword.Size = New System.Drawing.Size(195, 21)
        Me.txtUAAPassword.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 50)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Password"
        '
        'txtUAAUserName
        '
        Me.txtUAAUserName.Location = New System.Drawing.Point(76, 20)
        Me.txtUAAUserName.Name = "txtUAAUserName"
        Me.txtUAAUserName.Size = New System.Drawing.Size(195, 21)
        Me.txtUAAUserName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "User Name"
        '
        'cmdApply
        '
        Me.cmdApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdApply.Image = Global.LASERSystem.My.Resources.Resources.Save
        Me.cmdApply.Location = New System.Drawing.Point(658, 393)
        Me.cmdApply.Name = "cmdApply"
        Me.cmdApply.Size = New System.Drawing.Size(67, 28)
        Me.cmdApply.TabIndex = 7
        Me.cmdApply.Text = "Apply"
        Me.cmdApply.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdApply.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.Image = Global.LASERSystem.My.Resources.Resources.Done
        Me.cmdOK.Location = New System.Drawing.Point(585, 393)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(67, 28)
        Me.cmdOK.TabIndex = 6
        Me.cmdOK.Text = "OK"
        Me.cmdOK.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'cmbDBProvider
        '
        Me.cmbDBProvider.FormattingEnabled = True
        Me.cmbDBProvider.Items.AddRange(New Object() {"Microsoft.ACE.OLEDB.12.0"})
        Me.cmbDBProvider.Location = New System.Drawing.Point(110, 60)
        Me.cmbDBProvider.Name = "cmbDBProvider"
        Me.cmbDBProvider.Size = New System.Drawing.Size(215, 21)
        Me.cmbDBProvider.TabIndex = 9
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 63)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(98, 13)
        Me.Label10.TabIndex = 10
        Me.Label10.Text = "Database Provider:"
        '
        'FrmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(730, 433)
        Me.Controls.Add(Me.cmdApply)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.tcSettings)
        Me.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "LASER System - Settings"
        Me.tcSettings.ResumeLayout(False)
        Me.tpGeneral.ResumeLayout(False)
        Me.Guna2GroupBox4.ResumeLayout(False)
        Me.Guna2GroupBox4.PerformLayout()
        Me.Guna2GroupBox3.ResumeLayout(False)
        Me.Guna2GroupBox3.PerformLayout()
        Me.tpDatabase.ResumeLayout(False)
        Me.tpDatabase.PerformLayout()
        Me.tpPrinter.ResumeLayout(False)
        Me.Guna2GroupBox2.ResumeLayout(False)
        Me.Guna2GroupBox2.PerformLayout()
        Me.Guna2GroupBox1.ResumeLayout(False)
        Me.Guna2GroupBox1.PerformLayout()
        Me.tpUserAccount.ResumeLayout(False)
        CType(Me.grdUAUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpUAUser.ResumeLayout(False)
        Me.grpUAUser.PerformLayout()
        Me.grpUAAdmin.ResumeLayout(False)
        Me.grpUAAdmin.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tcSettings As TabControl
    Friend WithEvents tpDatabase As TabPage
    Friend WithEvents cmdDBLocation As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtDBLoc As TextBox
    Friend WithEvents cmdApply As Button
    Friend WithEvents cmdOK As Button
    Friend WithEvents ofdDatabase As OpenFileDialog
    Friend WithEvents tpGeneral As TabPage
    Friend WithEvents tpUserAccount As TabPage
    Friend WithEvents grpUAAdmin As GroupBox
    Friend WithEvents txtUAAPassword As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txtUAAUserName As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents cmdUACheck As Button
    Friend WithEvents grpUAUser As GroupBox
    Friend WithEvents cmdUADelete As Button
    Friend WithEvents cmdUASave As Button
    Friend WithEvents cmdUANew As Button
    Friend WithEvents txtUANewPW As TextBox
    Friend WithEvents lblUANewPW As Label
    Friend WithEvents txtUAUserName As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txtUAReenterPW As TextBox
    Friend WithEvents lblUAReenterPW As Label
    Friend WithEvents txtUAEmail As TextBox
    Friend WithEvents lblUAEmail As Label
    Friend WithEvents grdUAUser As DataGridView
    Friend WithEvents txtUACurrentPW As TextBox
    Friend WithEvents lblUACurrentPW As Label
    Friend WithEvents txtUAUNo As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents cmbUAType As ComboBox
    Friend WithEvents UAUNo As DataGridViewTextBoxColumn
    Friend WithEvents UAUserName As DataGridViewTextBoxColumn
    Friend WithEvents UAType As DataGridViewTextBoxColumn
    Friend WithEvents UAEmail As DataGridViewTextBoxColumn
    Friend WithEvents tpPrinter As TabPage
    Friend WithEvents Guna2GroupBox1 As Guna.UI2.WinForms.Guna2GroupBox
    Friend WithEvents txtStickerRepairPaperName As TextBox
    Friend WithEvents txtStickerPrinterName As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents Guna2GroupBox2 As Guna.UI2.WinForms.Guna2GroupBox
    Friend WithEvents txtBillPaperName As TextBox
    Friend WithEvents txtBillPrinterName As TextBox
    Friend WithEvents Label22 As Label
    Friend WithEvents Label23 As Label
    Friend WithEvents txtStickerStockPaperName As TextBox
    Friend WithEvents Label24 As Label
    Friend WithEvents Guna2GroupBox3 As Guna.UI2.WinForms.Guna2GroupBox
    Friend WithEvents txtMAdminEmail As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents chkDeliveredEmailtoT As CheckBox
    Friend WithEvents chkMSetEmail As CheckBox
    Friend WithEvents Guna2GroupBox4 As Guna.UI2.WinForms.Guna2GroupBox
    Friend WithEvents chkBSCOMMode As CheckBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtBSBaudRate As TextBox
    Friend WithEvents cmbBSCOMPort As ComboBox
    Friend WithEvents txtDBPassword As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents cmbDBProvider As ComboBox
End Class
