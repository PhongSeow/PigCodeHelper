<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.timMain = New System.Windows.Forms.Timer(Me.components)
        Me.btnEnumValue = New System.Windows.Forms.Button()
        Me.tbMain = New System.Windows.Forms.TextBox()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnClassValue = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'timMain
        '
        Me.timMain.Interval = 1000
        '
        'btnEnumValue
        '
        Me.btnEnumValue.Location = New System.Drawing.Point(61, 12)
        Me.btnEnumValue.Name = "btnEnumValue"
        Me.btnEnumValue.Size = New System.Drawing.Size(93, 23)
        Me.btnEnumValue.TabIndex = 1
        Me.btnEnumValue.Text = "获取枚举常量项"
        Me.btnEnumValue.UseVisualStyleBackColor = True
        '
        'tbMain
        '
        Me.tbMain.Location = New System.Drawing.Point(12, 82)
        Me.tbMain.Multiline = True
        Me.tbMain.Name = "tbMain"
        Me.tbMain.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.tbMain.Size = New System.Drawing.Size(373, 386)
        Me.tbMain.TabIndex = 2
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(12, 12)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(43, 23)
        Me.btnStop.TabIndex = 3
        Me.btnStop.Text = "结束"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnClassValue
        '
        Me.btnClassValue.Location = New System.Drawing.Point(160, 12)
        Me.btnClassValue.Name = "btnClassValue"
        Me.btnClassValue.Size = New System.Drawing.Size(93, 23)
        Me.btnClassValue.TabIndex = 4
        Me.btnClassValue.Text = "获取类接口"
        Me.btnClassValue.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(397, 480)
        Me.Controls.Add(Me.btnClassValue)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.tbMain)
        Me.Controls.Add(Me.btnEnumValue)
        Me.Name = "frmMain"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents timMain As Timer
    Friend WithEvents btnEnumValue As Button
    Friend WithEvents tbMain As TextBox
    Friend WithEvents btnStop As Button
    Friend WithEvents btnClassValue As Button
End Class
