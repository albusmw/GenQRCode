Option Explicit On
Option Strict On

Public Class MainForm

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Get build data
        Dim BuildDate As String = String.Empty
        Dim AllResources As String() = System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceNames
        For Each Entry As String In AllResources
            If Entry.EndsWith(".BuildDate.txt") Then
                BuildDate = " (Build of " & (New System.IO.StreamReader(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(Entry)).ReadToEnd.Trim).Replace(",", ".") & ")"
                Exit For
            End If
        Next Entry
        Me.Text &= BuildDate

        'Display code
        GenCode()

    End Sub

    Private Sub GenCode()

        Try
            Dim Generator As New QRCoder.PayloadGenerator.Girocode(tbIBAN.Text, tbBIC.Text, tbReceiver.Text, CType(tbAmount.Text, Decimal), tbPurpose.Text)
            Dim Code As New QRCoder.QRCode((New QRCoder.QRCodeGenerator).CreateQrCode(Generator.ToString, QRCoder.QRCodeGenerator.ECCLevel.M))
            pbQRCode.Image = Code.GetGraphic(CInt(tbPixelPerBit.Text))
            tbErrorText.Text = "---"
        Catch ex As Exception
            tbErrorText.Text = ex.Message
        End Try

    End Sub

    Private Sub tbIBAN_TextChanged(sender As Object, e As EventArgs) Handles tbIBAN.TextChanged, tbBIC.TextChanged, tbReceiver.TextChanged, tbAmount.TextChanged, tbPurpose.TextChanged, tbPixelPerBit.TextChanged
        GenCode()
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        Clipboard.Clear()
        Clipboard.SetImage(pbQRCode.Image)
    End Sub

    Private Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.KeyCode = Keys.C And e.Control = True Then
            btnCopy_Click(Nothing, Nothing)
            tbErrorText.Text = "Code copied to clipboard"
        End If
    End Sub
End Class
