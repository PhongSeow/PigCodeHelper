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
    Private oPigFunc As New PigFunc

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

    ''' <summary>
    ''' VB6对象浏览器到VB.NET的代码转换-类接口
    ''' </summary>
    ''' <param name="ConvWhat">转换什么</param>
    ''' <param name="SrcStr">源串</param>
    ''' <returns></returns>
    Public Function VB6ObjBrow2VBNetCode_ClassValue(ConvWhat As enmConvWhat, SrcStr As String) As String
        Try
            VB6ObjBrow2VBNetCode_ClassValue = ""
            SrcStr = Trim(SrcStr)
            Select Case ConvWhat
                Case enmConvWhat.Head
                    SrcStr = Replace(SrcStr, "Class ", "")
                    VB6ObjBrow2VBNetCode_ClassValue = "Public Class " & SrcStr & vbCrLf
                    VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Inherits PigBaseMini" & vbCrLf
                    VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Private Const CLS_VERSION As String = ""1.0.1""" & vbCrLf
                    VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Public Obj As Object" & vbCrLf
                    VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Public Sub New()" & vbCrLf
                    VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "MyBase.New(CLS_VERSION)" & vbCrLf
                    VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "End Sub" & vbCrLf
                Case enmConvWhat.Bottom
                    VB6ObjBrow2VBNetCode_ClassValue = "End Class" & vbCrLf
                Case enmConvWhat.BodyItem
                    Dim strData As String = SrcStr & vbCrLf, strHead As String = oPigFunc.GetStr(strData, "", " ")
                    Dim strSubName As String = "", strLeft As String = "", strRight As String = "", strDataType As String = ""
                    Dim bolIsReadOnly As Boolean = False, strPara As String = ""
                    Select Case strHead
                        Case "Property"
                            strSubName = oPigFunc.GetStr(strData, "", " ")
                            strDataType = oPigFunc.GetStr(strData, "As ", vbCrLf)
                            If InStr(strDataType, vbLf) > 0 Then
                                strDataType = oPigFunc.GetStr(strDataType, "", vbLf)
                            End If
                            strDataType = Replace(strDataType, vbCrLf, "")
                            If InStr(SrcStr, "read-only") > 0 Or InStr(SrcStr, "只读") > 0 Then bolIsReadOnly = True
                            VB6ObjBrow2VBNetCode_ClassValue = "Public "
                            If bolIsReadOnly = True Then VB6ObjBrow2VBNetCode_ClassValue &= " ReadOnly"
                            VB6ObjBrow2VBNetCode_ClassValue &= " Property " & strSubName & "() As " & strDataType & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Get" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Try" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & vbTab & "Return Me.Obj." & strSubName & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""" & strSubName & ".Get"", ex)" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & vbTab & "Return Nothing" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "End Try" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "End Get" & vbCrLf
                            If bolIsReadOnly = False Then
                                VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Set(value As " & strDataType & ")" & vbCrLf
                                VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Try" & vbCrLf
                                VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & vbTab & "Me.Obj." & strSubName & " = value" & vbCrLf
                                VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
                                VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
                                VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""" & strSubName & ".Set"", ex)" & vbCrLf
                                VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "End Try" & vbCrLf
                                VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "End Set" & vbCrLf
                            End If
                            VB6ObjBrow2VBNetCode_ClassValue &= "End Property" & vbCrLf
                        Case "Function"
                            If InStr(strData, "()") > 0 Then
                                strSubName = oPigFunc.GetStr(strData, "", "()")
                                strPara = ""
                            Else
                                strSubName = oPigFunc.GetStr(strData, "", "(")
                                strPara = oPigFunc.GetStr(strData, "", ")")
                                strPara = Replace(strPara, "[", "Optional ")
                                strPara = Replace(strPara, "]", " = """"")
                            End If
                            strDataType = oPigFunc.GetStr(strData, " As ", vbCrLf)
                            If InStr(strDataType, vbLf) > 0 Then
                                strDataType = oPigFunc.GetStr(strDataType, "", vbLf)
                            End If
                            strDataType = Replace(strDataType, vbCrLf, "")
                            VB6ObjBrow2VBNetCode_ClassValue = "Public Function " & strSubName & "(" & strPara & ") As " & strDataType & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Try" & vbCrLf
                            Do While True
                                If InStr(strPara, " As ") = 0 Then Exit Do
                                If Right(strPara, 1) <> " " Then strPara &= " "
                                Dim strTmp As String = oPigFunc.GetStr(strPara, " As ", " ")
                            Loop
                            strPara = Replace(strPara, "Optional", " ")
                            strPara = Replace(strPara, "= """"", " ")
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & strSubName & " = Me.Obj." & strSubName & "(" & strPara & ")" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Me.ClearErr()" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Catch ex As Exception" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Me.SetSubErrInf(""" & strSubName & """, ex)" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Return Nothing" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "End Try" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= "End Function" & vbCrLf
                        Case "Sub"
                            If InStr(strData, "()") > 0 Then
                                strSubName = oPigFunc.GetStr(strData, "", "()")
                                strPara = ""
                            Else
                                strSubName = oPigFunc.GetStr(strData, "", "(")
                                strPara = oPigFunc.GetStr(strData, "", ")")
                                strPara = Replace(strPara, "[", "Optional ")
                                strPara = Replace(strPara, "]", " = """"")
                            End If
                            VB6ObjBrow2VBNetCode_ClassValue = "Public Sub " & strSubName & "(" & strPara & ")" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Try" & vbCrLf
                            Do While True
                                If InStr(strPara, " As ") = 0 Then Exit Do
                                If Right(strPara, 1) <> " " Then strPara &= " "
                                Dim strTmp As String = oPigFunc.GetStr(strPara, " As ", " ")
                            Loop
                            strPara = Replace(strPara, "Optional", " ")
                            strPara = Replace(strPara, "= """"", " ")
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Me.Obj." & strSubName & "(" & strPara & ")" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Me.ClearErr()" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "Catch ex As Exception" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & vbTab & "Me.SetSubErrInf(""" & strSubName & """, ex)" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= vbTab & "End Try" & vbCrLf
                            VB6ObjBrow2VBNetCode_ClassValue &= "End Sub" & vbCrLf
                        Case "Event"
                    End Select
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("VB6ObjBrow2VBNetCode_ClassValue", ex)
            Return ""
        End Try
    End Function

End Class
