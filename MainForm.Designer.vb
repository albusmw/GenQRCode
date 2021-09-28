<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.tbIBAN = New System.Windows.Forms.TextBox()
        Me.tbBIC = New System.Windows.Forms.TextBox()
        Me.tbReceiver = New System.Windows.Forms.TextBox()
        Me.tbAmount = New System.Windows.Forms.TextBox()
        Me.pbQRCode = New System.Windows.Forms.PictureBox()
        Me.tbErrorText = New System.Windows.Forms.TextBox()
        Me.gbMain = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbPixelPerBit = New System.Windows.Forms.TextBox()
        Me.btnCopy = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbPurpose = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.scMain = New System.Windows.Forms.SplitContainer()
        CType(Me.pbQRCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbMain.SuspendLayout()
        CType(Me.scMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.scMain.Panel1.SuspendLayout()
        Me.scMain.Panel2.SuspendLayout()
        Me.scMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbIBAN
        '
        Me.tbIBAN.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbIBAN.Location = New System.Drawing.Point(115, 45)
        Me.tbIBAN.Name = "tbIBAN"
        Me.tbIBAN.Size = New System.Drawing.Size(208, 20)
        Me.tbIBAN.TabIndex = 0
        Me.tbIBAN.Text = "DE02100100100006820101"
        '
        'tbBIC
        '
        Me.tbBIC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbBIC.Location = New System.Drawing.Point(115, 71)
        Me.tbBIC.Name = "tbBIC"
        Me.tbBIC.Size = New System.Drawing.Size(208, 20)
        Me.tbBIC.TabIndex = 1
        Me.tbBIC.Text = "ABCDEFG1001"
        '
        'tbReceiver
        '
        Me.tbReceiver.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbReceiver.Location = New System.Drawing.Point(115, 19)
        Me.tbReceiver.Name = "tbReceiver"
        Me.tbReceiver.Size = New System.Drawing.Size(208, 20)
        Me.tbReceiver.TabIndex = 2
        Me.tbReceiver.Text = "Max Mustermann"
        '
        'tbAmount
        '
        Me.tbAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbAmount.Location = New System.Drawing.Point(115, 97)
        Me.tbAmount.Name = "tbAmount"
        Me.tbAmount.Size = New System.Drawing.Size(208, 20)
        Me.tbAmount.TabIndex = 3
        Me.tbAmount.Text = "484,49"
        '
        'pbQRCode
        '
        Me.pbQRCode.BackColor = System.Drawing.Color.Silver
        Me.pbQRCode.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pbQRCode.Location = New System.Drawing.Point(0, 0)
        Me.pbQRCode.Name = "pbQRCode"
        Me.pbQRCode.Size = New System.Drawing.Size(470, 352)
        Me.pbQRCode.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pbQRCode.TabIndex = 4
        Me.pbQRCode.TabStop = False
        '
        'tbErrorText
        '
        Me.tbErrorText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbErrorText.Location = New System.Drawing.Point(12, 370)
        Me.tbErrorText.Name = "tbErrorText"
        Me.tbErrorText.ReadOnly = True
        Me.tbErrorText.Size = New System.Drawing.Size(809, 20)
        Me.tbErrorText.TabIndex = 5
        '
        'gbMain
        '
        Me.gbMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbMain.Controls.Add(Me.Label6)
        Me.gbMain.Controls.Add(Me.tbPixelPerBit)
        Me.gbMain.Controls.Add(Me.btnCopy)
        Me.gbMain.Controls.Add(Me.Label5)
        Me.gbMain.Controls.Add(Me.tbPurpose)
        Me.gbMain.Controls.Add(Me.Label4)
        Me.gbMain.Controls.Add(Me.Label3)
        Me.gbMain.Controls.Add(Me.Label2)
        Me.gbMain.Controls.Add(Me.Label1)
        Me.gbMain.Controls.Add(Me.tbReceiver)
        Me.gbMain.Controls.Add(Me.tbAmount)
        Me.gbMain.Controls.Add(Me.tbIBAN)
        Me.gbMain.Controls.Add(Me.tbBIC)
        Me.gbMain.Location = New System.Drawing.Point(3, 3)
        Me.gbMain.Name = "gbMain"
        Me.gbMain.Size = New System.Drawing.Size(329, 346)
        Me.gbMain.TabIndex = 6
        Me.gbMain.TabStop = False
        Me.gbMain.Text = "Daten"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 152)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 13)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Pixel pro Bit"
        '
        'tbPixelPerBit
        '
        Me.tbPixelPerBit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPixelPerBit.Location = New System.Drawing.Point(115, 149)
        Me.tbPixelPerBit.Name = "tbPixelPerBit"
        Me.tbPixelPerBit.Size = New System.Drawing.Size(208, 20)
        Me.tbPixelPerBit.TabIndex = 8
        Me.tbPixelPerBit.Text = "10"
        '
        'btnCopy
        '
        Me.btnCopy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopy.Location = New System.Drawing.Point(9, 310)
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.Size = New System.Drawing.Size(314, 30)
        Me.btnCopy.TabIndex = 7
        Me.btnCopy.Text = "Bild ins Clipboard"
        Me.btnCopy.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 126)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(103, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Verwendungszweck"
        '
        'tbPurpose
        '
        Me.tbPurpose.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbPurpose.Location = New System.Drawing.Point(115, 123)
        Me.tbPurpose.Name = "tbPurpose"
        Me.tbPurpose.Size = New System.Drawing.Size(208, 20)
        Me.tbPurpose.TabIndex = 5
        Me.tbPurpose.Text = "Rechnung RE2020-0626"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 100)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(38, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Betrag"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 74)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "BIC"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Empfänger"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "IBAN"
        '
        'scMain
        '
        Me.scMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.scMain.Location = New System.Drawing.Point(12, 12)
        Me.scMain.Name = "scMain"
        '
        'scMain.Panel1
        '
        Me.scMain.Panel1.Controls.Add(Me.gbMain)
        '
        'scMain.Panel2
        '
        Me.scMain.Panel2.Controls.Add(Me.pbQRCode)
        Me.scMain.Size = New System.Drawing.Size(809, 352)
        Me.scMain.SplitterDistance = 335
        Me.scMain.TabIndex = 7
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(833, 402)
        Me.Controls.Add(Me.scMain)
        Me.Controls.Add(Me.tbErrorText)
        Me.KeyPreview = True
        Me.Name = "MainForm"
        Me.Text = "GenQRCode"
        CType(Me.pbQRCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbMain.ResumeLayout(False)
        Me.gbMain.PerformLayout()
        Me.scMain.Panel1.ResumeLayout(False)
        Me.scMain.Panel2.ResumeLayout(False)
        CType(Me.scMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.scMain.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tbIBAN As TextBox
    Friend WithEvents tbBIC As TextBox
    Friend WithEvents tbReceiver As TextBox
    Friend WithEvents tbAmount As TextBox
    Friend WithEvents pbQRCode As PictureBox
    Friend WithEvents tbErrorText As TextBox
    Friend WithEvents gbMain As GroupBox
    Friend WithEvents Label5 As Label
    Friend WithEvents tbPurpose As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents scMain As SplitContainer
    Friend WithEvents btnCopy As Button
    Friend WithEvents Label6 As Label
    Friend WithEvents tbPixelPerBit As TextBox
End Class
