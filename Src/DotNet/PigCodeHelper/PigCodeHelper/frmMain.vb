Imports PigCodeHelperLib
Public Class frmMain
    Public CurrText As String
    Public FullText As String
    Public IsHeadBegin As Boolean
    Public Enum emnDoWhat
        Unknow = 0
        GetEnumValue = 1    '获取枚举常量
    End Enum

    Public DoWhat As emnDoWhat
    Public HelperApp As New HelperApp

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Me
            .Text = "PigCodeHelper"
            .TopMost = True '置顶
            .MinimizeBox = False
            .MaximizeBox = False
            '显示在屏幕右上角
            .StartPosition = FormStartPosition.Manual
            .Location = New Point(SystemInformation.PrimaryMonitorSize.Width - Me.Width - 10, 10)
            '透明度
            .Opacity = 0.8
            .DoWhat = emnDoWhat.Unknow
        End With
    End Sub
    Public Sub SetText(TextCont As String)
        System.Windows.Forms.Clipboard.SetText(TextCont)
    End Sub

    Public Sub GetText()
        Dim strText As String = System.Windows.Forms.Clipboard.GetText()
        If strText <> "" Then
            Me.CurrText = strText
            System.Windows.Forms.Clipboard.Clear()
        End If
    End Sub


    Private Sub timMain_Tick(sender As Object, e As EventArgs) Handles timMain.Tick
        Select Case Me.DoWhat
            Case emnDoWhat.GetEnumValue
                Me.GetText()
                If Me.CurrText <> "" Then
                    If Mid(Me.CurrText, 1, 5) = "Enum " And Me.IsHeadBegin = False Then
                        Me.IsHeadBegin = True
                        Me.FullText = Me.HelperApp.VB6ObjBrow2VBNetCode_EnumValue(HelperApp.enmConvWhat.Head, Me.CurrText)
                    ElseIf Mid(Me.CurrText, 1, 6) = "Const " And Me.IsHeadBegin = True Then
                        Me.FullText &= Me.HelperApp.VB6ObjBrow2VBNetCode_EnumValue(HelperApp.enmConvWhat.BodyItem, Me.CurrText)
                    End If
                    Me.CurrText = ""
                    Me.tbMain.Text = Me.FullText
                End If
        End Select
    End Sub

    Public Sub BeginGet()
        System.Windows.Forms.Clipboard.Clear()
        Me.FullText = ""
        Me.CurrText = ""
        Me.IsHeadBegin = False
        Me.timMain.Enabled = True
    End Sub

    Private Sub btnEnumValue_Click(sender As Object, e As EventArgs) Handles btnEnumValue.Click
        Me.DoWhat = emnDoWhat.GetEnumValue
        Me.Text = "处理" & Me.btnEnumValue.Text
        Me.BeginGet()
    End Sub


    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        Select Case Me.DoWhat
            Case emnDoWhat.GetEnumValue
                Me.FullText &= Me.HelperApp.VB6ObjBrow2VBNetCode_EnumValue(HelperApp.enmConvWhat.Bottom, "")
        End Select
        Me.DoWhat = emnDoWhat.Unknow
        Me.timMain.Enabled = False
        Me.SetText(Me.FullText)
        Me.tbMain.Text = Me.FullText
    End Sub
End Class
