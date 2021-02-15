'**********************************
'* Name: HelperApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Helper Appication
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 13/2/2021
'**********************************
Imports PigToolsLib
Public Class HelperApp
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.1"


    Public Enum enmConvWhat
        Head = 10
        BodyItem = 20
        Bottom = 30
    End Enum

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    ''' <summary>
    ''' VB6对象浏览器到VB.NET的代码转换-枚举定义
    ''' </summary>
    ''' <param name="ConvWhat">转换什么</param>
    ''' <param name="SrcStr">源串</param>
    ''' <returns></returns>
    Public Function VB6ObjBrow2VBNetCode_EnumValue(ConvWhat As enmConvWhat, SrcStr As String) As String
        Try
            VB6ObjBrow2VBNetCode_EnumValue = ""
            SrcStr = Trim(SrcStr)
            Select Case ConvWhat
                Case enmConvWhat.Head
                    SrcStr = Replace(SrcStr, "Enum ", "")
                    VB6ObjBrow2VBNetCode_EnumValue = "Public Enum " & SrcStr & vbCrLf
                Case enmConvWhat.Bottom
                    VB6ObjBrow2VBNetCode_EnumValue = "End Enum" & vbCrLf
                Case enmConvWhat.BodyItem
                    SrcStr = Replace(SrcStr, "Const ", "")
                    VB6ObjBrow2VBNetCode_EnumValue = vbTab & SrcStr & vbCrLf
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("VB6ObjBrow2VBNetCode_EnumValue", ex)
            Return ""
        End Try
    End Function

End Class
