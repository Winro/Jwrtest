Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class ProductionMaterials
    Private LH_Grinding_material_number As String
    Private LH_Material_type As String
    Private LH_Name_specifications As String
    Private SqlConn As New SqlConnection(ConfigurationManager.ConnectionStrings("NB410Sys").ConnectionString)
    Private Sqlcmd As SqlCommand = SqlConn.CreateCommand
    Private Dr As SqlDataReader
    Private Trans As SqlTransaction



    Public Sub New(ByRef MaterialName As String)
        LH_Grinding_material_number = MaterialName
        LH_Material_type = getLH_Material_type()
        LH_Name_specifications = getLH_Name_specifications()
    End Sub
    Public Sub New(ByRef MaterialName As String, ByRef MaterialSpecifications As String)
        LH_Grinding_material_number = MaterialName
        LH_Material_type = getLH_Material_type()
        LH_Name_specifications = MaterialSpecifications.Replace(",", ";")
    End Sub
    Public Function ImportLH(ByRef UpdatePerson As String) As String
        Return CheckLHExist(UpdatePerson)
    End Function
    ''' <summary>
    ''' 导入料号后判断是否创建或者更新原料号
    ''' </summary>
    ''' <param name="UpdatePerson">更新人员</param>
    ''' <returns></returns>
    Private Function CheckLHExist(ByRef UpdatePerson As String) As String
        If SqlSearch_NB_Lh_Basic("LH_Grinding_material_number") = "" Then
            If LH_Name_specifications = "" Then
                Return "品名规格不允许为空值！"
            Else
                Return SqlInsert_NB_Lh_Basic(UpdatePerson)
            End If
        Else
            If LH_Name_specifications = SqlSearch_NB_Lh_Basic("LH_Name_specifications") Then
                Return "内容相同，未更新"
            Else
                Return SqlUpdate_NB_Lh_Basic(UpdatePerson)
            End If
        End If
    End Function

    Private Function SqlUpdate_NB_Lh_Basic(ByRef UpdatePerson As String) As String
        Dim Sql_Str As String = ""
        Sql_Str = Sql_Str & "Insert into NB_LH_io (LH_Grinding_material_number, LH_Name_specifications_old"
        Sql_Str = Sql_Str & " , LH_Name_specifications_new, LH_update_time, LH_update_person) Values"
        Sql_Str = Sql_Str & " ('" + LH_Grinding_material_number + "','" + SqlSearch_NB_Lh_Basic("LH_Name_specifications") + "'"
        Sql_Str = Sql_Str & " ,'" + LH_Name_specifications + "','" + Now + "','" + UpdatePerson + "')"
        Sql_Str = Sql_Str & chr(13)
        Sql_Str = Sql_Str & " Update NB_Lh_Basic Set "
        Sql_Str = Sql_Str & " LH_Name_specifications='" + LH_Name_specifications + "'"
        Sql_Str = Sql_Str & " ,LH_Material_code='" + getLH_Material_code() + "'"
        Sql_Str = Sql_Str & " ,LH_Material_type='" + getLH_Material_type() + "'"
        Sql_Str = Sql_Str & " ,LH_size_Film_thickness='" & getLH_size_Film_thickness() & "'"
        Sql_Str = Sql_Str & " ,LH_Adhesive='" + getLH_Adhesive() + "'"
        Sql_Str = Sql_Str & " ,LH_texture_of_material='" + getLH_texture_of_material() + "'"
        Select Case LH_Material_type
            Case "S", "U"
                Sql_Str = Sql_Str & " ,Material_breadth='" & getMaterial_breadth() & "'"
                Sql_Str = Sql_Str & " ,Material_type='" + getMaterial_type() + "'"
            Case Else
                Sql_Str = Sql_Str & " ,Lh_Size='" + getLh_Size() + "'"
                Sql_Str = Sql_Str & " ,LH_Size_coding='" + getLH_Size_coding() + "'"
                Sql_Str = Sql_Str & " ,LH_size_length='" & getLH_size_length() & "'"
                Sql_Str = Sql_Str & " ,LH_size_wide='" & getLH_size_wide() & "'"
                Sql_Str = Sql_Str & " ,LH_stamp='" + getLH_stamp() + "'"
                Sql_Str = Sql_Str & " ,LH_stamp_Position='" + getLH_stamp_Position() + "'"
                If getLH_angle_inf() = "" Then
                    Sql_Str = Sql_Str & " ,LH_angle_inf=NULL"
                    Sql_Str = Sql_Str & " ,A_angle_size=NULL"
                    Sql_Str = Sql_Str & " ,B_angle_size=NULL"
                    Sql_Str = Sql_Str & " ,C_angle_size=NULL"
                    Sql_Str = Sql_Str & " ,D_angle_size=NULL"
                Else
                    Sql_Str = Sql_Str & " ,LH_angle_inf='" + getLH_angle_inf() + "'"
                    Sql_Str = Sql_Str & " ,A_angle_size='" & getAngle_size("A") & "'"
                    Sql_Str = Sql_Str & " ,B_angle_size='" & getAngle_size("B") & "'"
                    Sql_Str = Sql_Str & " ,C_angle_size='" & getAngle_size("C") & "'"
                    Sql_Str = Sql_Str & " ,D_angle_size='" & getAngle_size("D") & "'"
                End If
                Sql_Str = Sql_Str & " ,LH_size_thickness='" & getLH_size_thickness() & "'"
                Sql_Str = Sql_Str & " ,LH_Size_category='" + getLH_Size_category() + "'"
                Sql_Str = Sql_Str & " ,LH_size_thickness_type='" + getLH_size_thickness_type() + "'"
        End Select
        Sql_Str = Sql_Str & " ,LH_update_time='" + now + "'"
        Sql_Str = Sql_Str & " ,LH_update_person='" + UpdatePerson + "'"
        Sql_Str = Sql_Str & " where LH_Grinding_material_number='" + LH_Grinding_material_number + "'"
        Return SqlRun(Sql_Str, 2)
    End Function
    Private Function SqlInsert_NB_Lh_Basic(ByRef UpdatePerson As String) As String
        Dim Sql_Str As String = ""
        Sql_Str = Sql_Str & "Insert into NB_Lh_Basic "
        Sql_Str = Sql_Str & " ( LH_Grinding_material_number"
        Sql_Str = Sql_Str & " ,LH_Name_specifications"
        Sql_Str = Sql_Str & " ,LH_Material_code"
        Sql_Str = Sql_Str & " ,LH_Material_type"
        Sql_Str = Sql_Str & " ,LH_size_Film_thickness"
        Sql_Str = Sql_Str & " ,LH_Adhesive"
        Sql_Str = Sql_Str & " ,LH_texture_of_material"
        Select Case LH_Material_type
            Case "S", "U"
                Sql_Str = Sql_Str & " ,Material_breadth"
                Sql_Str = Sql_Str & " ,Material_type"
            Case Else
                Sql_Str = Sql_Str & " ,Lh_Size"
                Sql_Str = Sql_Str & " ,LH_Size_coding"
                Sql_Str = Sql_Str & " ,LH_size_length"
                Sql_Str = Sql_Str & " ,LH_size_wide"
                Sql_Str = Sql_Str & " ,LH_stamp"
                Sql_Str = Sql_Str & " ,LH_stamp_Position"
                If Not getLH_angle_inf() = "" Then
                    Sql_Str = Sql_Str & " ,LH_angle_inf"
                    Sql_Str = Sql_Str & " ,A_angle_size"
                    Sql_Str = Sql_Str & " ,B_angle_size"
                    Sql_Str = Sql_Str & " ,C_angle_size"
                    Sql_Str = Sql_Str & " ,D_angle_size"
                End If
                Sql_Str = Sql_Str & " ,LH_size_thickness"
                Sql_Str = Sql_Str & " ,LH_Size_category"
                Sql_Str = Sql_Str & " ,LH_size_thickness_type"
        End Select
        Sql_Str = Sql_Str & " ,LH_update_time"
        Sql_Str = Sql_Str & " ,LH_update_person) Values "
        Sql_Str = Sql_Str & " ( '" + LH_Grinding_material_number + "'"
        Sql_Str = Sql_Str & " ,'" + LH_Name_specifications + "'"
        Sql_Str = Sql_Str & " ,'" + getLH_Material_code() + "'"
        Sql_Str = Sql_Str & " ,'" + getLH_Material_type() + "'"
        Sql_Str = Sql_Str & " ,'" & getLH_size_Film_thickness() & "'"
        Sql_Str = Sql_Str & " ,'" + getLH_Adhesive() + "'"
        Sql_Str = Sql_Str & " ,'" + getLH_texture_of_material() + "'"
        Select Case LH_Material_type
            Case "S", "U"
                Sql_Str = Sql_Str & " ,'" & getMaterial_breadth() & "'"
                Sql_Str = Sql_Str & " ,'" + getMaterial_type() + "'"
            Case Else
                Sql_Str = Sql_Str & " ,'" + getLh_Size() + "'"
                Sql_Str = Sql_Str & " ,'" + getLH_Size_coding() + "'"
                Sql_Str = Sql_Str & " ,'" & getLH_size_length() & "'"
                Sql_Str = Sql_Str & " ,'" & getLH_size_wide() & "'"
                Sql_Str = Sql_Str & " ,'" + getLH_stamp() + "'"
                Sql_Str = Sql_Str & " ,'" + getLH_stamp_Position() + "'"
                If Not getLH_angle_inf() = "" Then
                    Sql_Str = Sql_Str & " ,'" + getLH_angle_inf() + "'"
                    Sql_Str = Sql_Str & " ,'" & getAngle_size("A") & "'"
                    Sql_Str = Sql_Str & " ,'" & getAngle_size("B") & "'"
                    Sql_Str = Sql_Str & " ,'" & getAngle_size("C") & "'"
                    Sql_Str = Sql_Str & " ,'" & getAngle_size("D") & "'"
                End If
                Sql_Str = Sql_Str & " ,'" & getLH_size_thickness() & "'"
                Sql_Str = Sql_Str & " ,'" + getLH_Size_category() + "'"
                Sql_Str = Sql_Str & " ,'" + getLH_size_thickness_type+"'"
        End Select
        Sql_Str = Sql_Str & " ,'" + now + "'"
        Sql_Str = Sql_Str & " ,'" + UpdatePerson + "')"
        Return SqlRun（Sql_Str, 1）
    End Function
    Private Function SqlRun(ByRef SqlStr As String, ByRef item As Byte) As String
        SqlConn.Open()
        Trans = SqlConn.BeginTransaction '第一步，开始事务。这句必须在下句前，相当于实例事务对象
        Sqlcmd.Transaction = Trans     'Sqlcmd要做的事
        Try
            Dim i As Integer = 0
            Sqlcmd.CommandText = SqlStr
            i = Sqlcmd.ExecuteNonQuery()
            Trans.Commit()
            SqlConn.Close()
            Return "提交成功! " & i & " / " & item '& SqlStr
        Catch ex As Exception
            Trans.Rollback()
            SqlConn.Close()
            Return "提交失败: " & ex.Message.ToString & vbCrLf & SqlStr
        End Try
    End Function
    ''' <summary>
    ''' 获取品名规格
    ''' </summary>
    ''' <returns></returns>
    Function getLH_Name_specifications() As String
        Return SqlSearch_NB_Lh_Basic("LH_Name_specifications")
    End Function
    Private Function SqlSearch_NB_Lh_Basic(ByRef Item As String) As String
        SqlSearch_NB_Lh_Basic = ""
        Sqlcmd.CommandText = "SELECT  " + Item + "  FROM NB_Lh_Basic Where LH_Grinding_material_number='" + LH_Grinding_material_number + "'"
        SqlConn.Open()
        Try
            Dr = Sqlcmd.ExecuteReader
            If Dr.HasRows Then
                If Dr.Read Then
                    SqlSearch_NB_Lh_Basic = Dr(Item).ToString
                    Dr.Close()
                    SqlConn.Close()
                End If
            End If
            Dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If Dr IsNot Nothing AndAlso Not Dr.IsClosed Then Dr.Close()
            SqlConn.Close()
            SqlSearch_NB_Lh_Basic = "error: " & ex.Message & "  " & Sqlcmd.CommandText.ToString
        End Try
        Return SqlSearch_NB_Lh_Basic
    End Function
    ''' <summary>
    ''' 获取物料类型
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_Material_type() As String
        Return Left(LH_Grinding_material_number, 1)
    End Function
    ''' <summary>
    ''' 获取尺寸编码
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_Size_coding() As String
        Select Case LH_Material_type
            Case "S", "U"
                Return "-"
            Case "F", "I", "O"
                Return Mid$(LH_Grinding_material_number, 5, 3)
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' 获取物料编码
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_Material_code() As String
        Select Case LH_Material_type
            Case "S", "U"
                Return Right(LH_Grinding_material_number, 3)
            Case "F", "I", "O"
                Return Mid$(LH_Grinding_material_number, 8, 3)
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' 获取料号尺寸 
    ''' </summary>
    ''' <returns></returns>
    Private Function getLh_Size() As String
        Select Case LH_Material_type
            Case "S", "U"
                Return ""
            Case "F", "O"
                If InStr(LH_Name_specifications, Chr(34)) > 0 Then
                    Dim Str As String = Split(LH_Name_specifications, Chr(34))(0)
                    Return Right(Str, Len(Str) - InStrRev(Str, ";")) & Chr(34)
                ElseIf InStr(LH_Name_specifications, "片") > 0 Then
                    Return Mid(LH_Name_specifications, InStr(LH_Name_specifications, "片") - 2, 3)
                Else
                    Return ""
                End If
            Case "I"
                If InStr(LH_Name_specifications, Chr(34)) > 0 Then
                    Dim Str As String = Split(LH_Name_specifications, Chr(34))(0)
                    Return "中继片-" & Right(Str, Len(Str) - InStrRev(Str, ";")) & Chr(34)
                ElseIf InStr(LH_Name_specifications, "片") > 0 Then
                    Return Mid(LH_Name_specifications, InStr(LH_Name_specifications, "片") - 2, 3)
                Else
                    Return ""
                End If
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' 获取料号尺寸长
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_size_length() As Double
        Select Case LH_Material_type
            Case "S", "U"
                Return 0
            Case "F", "I", "O"
                Return Mid(LH_Name_specifications, InStrRev(Left(LH_Name_specifications, InStr(LH_Name_specifications, "*")), "(") + 1, InStr(LH_Name_specifications, "*") - InStrRev(Left(LH_Name_specifications, InStr(LH_Name_specifications, "*")), "(") - 1)
            Case Else
                Return 0
        End Select
    End Function
    ''' <summary>
    ''' 获取料号尺寸宽
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_size_wide() As Double
        Select Case LH_Material_type
            Case "S", "U"
                Return 0
            Case "F", "I", "O"
                Return Mid(LH_Name_specifications, InStr(LH_Name_specifications, "*") + 1, InStr(Mid(LH_Name_specifications, InStr(LH_Name_specifications, "*")), ")") - 2)
            Case Else
                Return 0
        End Select
    End Function
    ''' <summary>
    ''' 获取料号章型
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_stamp() As String
        Try
            Select Case LH_Material_type
                Case "S", "U"
                    Return ""
                Case "F", "I", "O"
                    If InStr(LH_Name_specifications, "矢印") > 0 Then
                        Dim Str As String = Split(Split(LH_Name_specifications, "矢印")(1), ";")(0)
                        If InStr(Str, "角") > 0 Then
                            If InStr(Split(Str, "/")(0), "角") > 0 Then
                                Return Split(Str, "/")(1)
                            ElseIf InStr(Split(Str, "/")(1), "角") > 0 Then
                                Return Split(Str, "/")(0)
                            Else
                                Return Str
                            End If
                        Else
                            Return ""
                        End If
                    Else
                        Return ""
                    End If
                Case Else
                    Return ""
            End Select
        Catch ex As Exception
            Return LH_Grinding_material_number & "; " & ex.Message.ToString
        End Try
    End Function
    ''' <summary>
    ''' 获取料号矢印章位置
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_stamp_Position() As String
        Try
            Select Case LH_Material_type
                Case "S", "U"
                    Return ""
                Case "F", "I", "O"
                    If InStr(LH_Name_specifications, "矢印") > 0 Then
                        Dim Str As String = Split(Split(LH_Name_specifications, "矢印")(1), ";")(0)
                        If InStr(Str, "角") > 0 Then
                            If InStr(Split(Str, "/")(0), "角") > 0 Then
                                Return Split(Str, "/")(0)
                            ElseIf InStr(Split(Str, "/")(1), "角") > 0 Then
                                Return Split(Str, "/")(1)
                            Else
                                Return Str
                            End If
                            Return ""
                        End If
                    Else
                        Return ""
                    End If
                Case Else
                    Return ""
            End Select
        Catch ex As Exception
            Return LH_Grinding_material_number & "; " & ex.Message.ToString
        End Try
    End Function
    ''' <summary>
    ''' 获取导角信息
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_angle_inf() As String
        Select Case LH_Material_type
            Case "S", "U"
                Return ""
            Case "F", "I", "O"
                If InStr(LH_Name_specifications, "導角") > 0 Then
                    Return Split(Split(LH_Name_specifications, "導角")(1), ";")(0)
                Else
                    Return ""
                End If
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' 获取导角大小
    ''' </summary>
    ''' <param name="Point">导角位置A\B\C\D</param>
    ''' <returns></returns>
    Private Function getAngle_size(ByRef Point As String) As Double
        Select Case LH_Material_type
            Case "S", "U"
                Return 0
            Case "F", "I", "O"
                If InStr(LH_Name_specifications, "導角") > 0 Then
                    Dim Str As String = Split(Split(LH_Name_specifications, "導角")(1), ";")(0)
                    If InStr(Str, "C0") > 0 Then
                        Str = Str.Replace("C0", "0")
                    ElseIf InStr(Str, "C1") > 0 Then
                        Str = Str.Replace("C1", "1")
                    Else
                        Return 0
                    End If
                    If Str Like "*" & Point & "*" Then
                        Str = Split(Str, Point)(1)
                        Dim s As Byte = 0 '状态标号
                        Dim Size As String = 0 '暂存尺寸大小
                        'P06103T7(AP1)(218.5*137.4);FP051-X101;10.1"W;78μ;矢印D/B角;導角A/B/C/D角各C0.3mm;230°
                        '针对EW8(478.16 * 300.2);00004;22"W;215μ;矢印J/D角;導角C角/C0.5//A角/C1
                        'P021AG203T7-EWE128T1(AP1)(164.4*123.9);PF021-A080;8";295μ;矢印F/B角;導角B角/C0.5mm;219°
                        'HT31HC1T1(AP1)(1000*600);00003;梯形片(下窄);140μ;矢印S1/D角;導角C角
                        For i As Byte = 1 To Len(Str)
                            If IsNumeric(Mid(Str, i, 1)) Or Mid(Str, i, 1) = "." Then
                                Size = Size & Mid(Str, i, 1)
                                s = 1
                            ElseIf Mid(Str, i, 1) = "m" Or (s = 1 And Mid(Str, i, 1) = "/") Then
                                Exit For
                            End If
                        Next
                        Return Format(Val(Size), "0.00")
                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            Case Else
                Return 0
        End Select
    End Function
    ''' <summary>
    ''' 获取撕膜后厚度
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_size_thickness() As Integer
        Select Case LH_Material_type
            Case "S", "U"
                Return 0
            Case "F", "I", "O"
                If InStr(LH_Name_specifications, "μ") > 0 Then
                    Dim Str As String = Split(LH_Name_specifications, "μ")(0)
                    Str = Mid(Str, InStrRev(Str, ";") + 1, Len(Str) - InStrRev(Str, ";"))
                    Return Str
                Else
                    Return 0
                End If
            Case Else
                Return 0
        End Select
    End Function
    ''' <summary>
    ''' 获取撕膜前厚度
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_size_Film_thickness() As Integer
        Select Case LH_Material_type
            Case "S", "U"
                If InStr(LH_Name_specifications, "μ") > 0 Then
                    Dim Str As String = Split(LH_Name_specifications, "μ")(0)
                    Str = Mid(Str, InStrRev(Str, ";") + 1, Len(Str) - InStrRev(Str, ";"))
                    Return Str
                Else
                    Return 0
                End If
            Case "F", "I", "O"
                Sqlcmd.CommandText = "SELECT  LH_size_Film_thickness  FROM NB_Lh_Basic Where LH_Grinding_material_number  Like 's____0" + getLH_Material_code() + "'"
                If SqlConn.State.open Then SqlConn.Open()
                Try
                    Dr = Sqlcmd.ExecuteReader
                    If Dr.HasRows Then
                        If Dr.Read Then
                            Dim Str As String = Dr("LH_size_Film_thickness").ToString
                            Dr.Close()
                            SqlConn.Close()
                            Return Str
                        Else
                            Dr.Close()
                            SqlConn.Close()
                            Return 0
                        End If
                    Else
                        Dr.Close()
                        SqlConn.Close()
                        Return 0
                    End If
                    Dr.Close()
                    SqlConn.Close()
                Catch ex As Exception
                    If Dr IsNot Nothing AndAlso Not Dr.IsClosed Then Dr.Close()
                    SqlConn.Close()
                    Return 0
                    'Return "error: " & ex.Message & "  " & Sqlcmd.CommandText.ToString
                End Try
            Case Else
                Return 0
        End Select
    End Function
    ''' <summary>
    ''' 获取胶系
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_Adhesive() As String
        Select Case LH_Material_type
            Case "S"
                Dim Str As String = ""
                If InStr(LH_Name_specifications, "HT") > 0 Then
                    Str = Split(LH_Name_specifications, "HT")(1)
                    If InStr(Str, "T") > 0 Then
                        Str = Mid(Str, InStr(Str, "T"), 2)
                    Else
                        Str = ""
                    End If
                ElseIf InStr(LH_Name_specifications, "FT") > 0 Then
                    Str = Split(LH_Name_specifications, "FT")(1)
                    If InStr(Str, "T") > 0 Then
                        Str = Mid(Str, InStr(Str, "T"), 2)
                    Else
                        Str = ""
                    End If
                ElseIf InStr(LH_Name_specifications, "TK") > 0 Then
                    Str = Split(LH_Name_specifications, "TK")(1)
                    If InStr(Str, "T") > 0 Then
                        Str = Mid(Str, InStr(Str, "T"), 2)
                    Else
                        Str = ""
                    End If
                Else
                    If InStrRev(LH_Name_specifications, "T") > 0 Then
                        Str = Mid(LH_Name_specifications, InStrRev(LH_Name_specifications, "T"), 2)
                    Else
                        Str = ""
                    End If
                End If
                Return Str
            'Case "F"
            '    Dim Str As String = Split(LH_Name_specifications, ";")(0)
            '    If Mid(LH_Grinding_material_number, 4, 1) = "B" Then '增亮膜类胶系不准确
            '        Str = ""
            '    Else
            '        If InStr(LH_Name_specifications, "HT") > 0 Then
            '            Str = Split(LH_Name_specifications, "HT")(1)
            '            If InStr(Str, "T") > 0 Then
            '                Str = Mid(Str, InStr(Str, "T"), 2)
            '            Else
            '                Str = ""
            '            End If
            '        ElseIf InStr(LH_Name_specifications, "FT") > 0 Then
            '            Str = Split(LH_Name_specifications, "FT")(1)
            '            If InStr(Str, "T") > 0 Then
            '                Str = Mid(Str, InStr(Str, "T"), 2)
            '            Else
            '                Str = ""
            '            End If
            '        ElseIf InStr(LH_Name_specifications, "TK") > 0 Then
            '            Str = Split(LH_Name_specifications, "TK")(1)
            '            If InStr(Str, "T") > 0 Then
            '                Str = Mid(Str, InStr(Str, "T"), 2)
            '            Else
            '                Str = ""
            '            End If
            '        Else
            '            If InStr(LH_Name_specifications, "T") > 0 Then
            '                Str = Mid(LH_Name_specifications, InStr(LH_Name_specifications, "T"), 2)
            '            Else
            '                Str = ""
            '            End If
            '        End If
            '    End If
            '    Return Str
            Case "I", "O", "F"
                Dim Str As String = Split(LH_Name_specifications, ";")(0)
                If InStrRev(LH_Name_specifications, "T") > 0 Then
                    Str = Mid(LH_Name_specifications, InStrRev(LH_Name_specifications, "T"), 2)
                Else
                    Str = ""
                End If
                Return Str
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' 获得材质
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_texture_of_material() As String
        Select Case LH_Material_type
            Case "S", "U", "F"
                Select Case Left(LH_Name_specifications, 2)
                    Case "EW"
                        Return "EWE"
                    Case "HT"
                        Return "HT"
                    Case "VW"
                        Return "VW"
                    Case Else
                        Return ""
                End Select
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' 获取尺寸别
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_Size_category() As String
        Select Case LH_Material_type
            Case "F", "I", "O"
                If IsNumeric(getLh_Size().Replace(Chr(34), "")) Then
                    If Val(getLh_Size()) > 11.6 Then
                        Return "大尺"
                    Else
                        Return "小尺"
                    End If
                Else
                    Return getLh_Size()
                End If
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' 获取厚度别
    ''' </summary>
    ''' <returns></returns>
    Private Function getLH_size_thickness_type() As String
        Select Case LH_Material_type
            Case "F", "S", "O", "U", "I"
                Dim Height As Integer = getLH_size_Film_thickness()
                If Height = 0 Then
                    Return ""
                Else
                    If Height > 300 Then
                        Return "普通"
                    ElseIf Height < 200 Then
                        Return "薄型"
                    Else
                        Return ""
                    End If
                End If
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' 获取卷料幅宽
    ''' </summary>
    ''' <returns></returns>
    Private Function getMaterial_breadth() As Double
        Select Case LH_Material_type
            Case "S", "U"
                Dim Str As String = ""
                If InStr(LH_Name_specifications, "mm") > 0 Then
                    Str = Split(LH_Name_specifications, "mm")(0)
                    Return Format(Val(Mid(Str, InStrRev(Str, ";") + 1, Len(Str) - InStrRev(Str, ";"))), "0.00")
                Else
                    Return ""
                End If
                Return Str
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' 卷料类别
    ''' </summary>
    ''' <returns></returns>
    Private Function getMaterial_type() As String
        Select Case LH_Material_type
            Case "S", "U"
                Return Mid$(LH_Grinding_material_number, 6, 1)
            Case Else
                Return ""
        End Select
    End Function
End Class
