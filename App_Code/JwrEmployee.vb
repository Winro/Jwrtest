Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Public Class JwrEmployee
    Dim SqlConn As New SqlConnection(ConfigurationManager.ConnectionStrings("NB410Sys").ConnectionString)
    Dim Sqlcmd As SqlCommand = SqlConn.CreateCommand
    Dim Trans As SqlTransaction
    Dim dr As SqlDataReader
    Private DataBase As String = " NB_NB410hr_data "
    Private SqlStr As String = ""
    Protected No As String
    Private Inf() As String
    Protected Creater As String
    Public Sub New()
    End Sub
    ''' <summary>
    ''' 功能说明：按工号创建对象
    ''' </summary>
    ''' <param name="value">工号或者姓名</param>
    Public Sub New(ByVal value As String)
        Dim Str As String = value.Trim
        If isGrapheme(Left(Str, 1)) Then
            No = Str.ToUpper
        Else
            No = SearchMsg("No", Str)
        End If
    End Sub
    ''' <summary>
    ''' 功能说明：按工号创建对象
    ''' </summary>
    ''' <param name="value">工号或者姓名</param>
    ''' <param name="Creater1">创建人</param>
    Public Sub New(ByVal value As String, ByRef Creater1 As String)
        Dim Str As String = value.Trim
        If isGrapheme(Left(Str, 1)) Then
            No = Str.ToUpper
        Else
            No = SearchMsg("No", Str)
        End If
        Creater = Creater1
    End Sub
    ''' <summary>
    ''' 功能说明：按工号，数组创建对象
    ''' </summary>
    ''' <param name="Inf1">数组：部门，工号,姓名，入职日期，联系方式</param>
    ''' <param name="Creater1">创建人</param>
    Public Sub New(ByVal Inf1() As String, ByRef Creater1 As String)
        No = Inf1(1).ToUpper
        Inf = Inf1
        Creater = Creater1
    End Sub
    ''' <summary>
    ''' 属性：基本信息
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property basicInformation As String
        Get
            Return Department & " " & Station & " " & Class_1 & " " & No & " " & Name & " " & Phone & " " & Attendance
        End Get
    End Property
    ''' <summary>
    ''' 属性：姓名
    ''' </summary>
    ''' <returns></returns>
    Property Name As String
        Get
            Return SearchMsg("Nm")
        End Get
        Set(value As String)
            'UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性：部门
    ''' </summary>
    ''' <returns></returns>
    Property Department As String
        Get
            Return SearchMsg("Department_Type")
        End Get
        Set(value As String)
            'UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性：站别
    ''' </summary>
    ''' <returns></returns>
    Property Station As String
        Get
            Return SearchMsg("Station")
        End Get
        Set(value As String)
            'UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性：班别
    ''' </summary>
    ''' <returns></returns>
    Property Class_1 As String
        Get
            Return SearchMsg("Class_1")
        End Get
        Set(value As String)
            'UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性：职务
    ''' </summary>
    ''' <returns></returns>
    Property Duty As String
        Get
            Return SearchMsg("Duty")
        End Get
        Set(value As String)
            'UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性：认证等级
    ''' </summary>
    ''' <returns></returns>
    Property Grade As String
        Get
            Return SearchMsg("Grade")
        End Get
        Set(value As String)
            'UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性：出勤状况
    ''' </summary>
    ''' <returns></returns>
    Property Attendance As String
        Get
            Return SearchMsg("State_of_Attendance")
        End Get
        Set(value As String)
            'UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性：联系方式
    ''' </summary>
    ''' <returns></returns>
    Property Phone As String
        Get
            Return SearchMsg("Phone_number")
        End Get
        Set(value As String)
            'UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性：权限
    ''' </summary>
    ''' <returns></returns>
    Property Purview As String
        Get
            Return searchPurview("AddUser", Duty)
        End Get
        Set(value As String)
            'UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    ''' <summary>
    ''' 功能说明：判断单字是否是字母
    ''' </summary>
    ''' <param name="Item"></param>
    ''' <returns></returns>
    Private Function isGrapheme(ByRef Item As String) As Boolean
        If (Asc(Item) >= 65 And Asc(Item) <= 90) OrElse (Asc(Item) >= 97 And Asc(Item) <= 122) Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 功能说明：判断部门编码是否正确
    ''' </summary>
    ''' <param name="Str"></param>
    ''' <returns></returns>
    Protected Function isDepartment(ByRef Str As String) As Boolean
        If Str Like "??###" Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 功能说明：判断工号是否正确
    ''' </summary>
    ''' <returns></returns>
    Private Function isNo(ByRef Str As String) As Boolean
        If Len(Str) = 6 AndAlso (Left(Str, 1) = "N" Or Left(Str, 1) = "N") Then
            Return True
        ElseIf Len(Str) = 7 AndAlso Left(Str, 2) = "NF" Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 功能说明：判断名字是否正确
    ''' </summary>
    ''' <param name="Str">姓名</param>
    ''' <returns></returns>
    Private Function isName(ByRef Str As String) As Boolean
        If Len(Str) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 功能说明：判断手机号是否正确
    ''' </summary>
    ''' <param name="Str">手机号</param>
    ''' <returns></returns>
    Private Function isPhone(ByRef Str As String) As Boolean
        If Len(Str) > 0 And Len(Str) <= 12 Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 功能说明：卡控创建新人员信息
    ''' </summary>
    ''' <returns></returns>
    Private Function creatJudge() As String
        If Not isDepartment(Inf(0)) Then Return "创建人员数组错误：部门"
        If Not isNo(Inf(1)) Then Return "创建人员数组错误：工号"
        If Not isName(Inf(2)) Then Return "创建人员数组错误：姓名"
        If Not IsDate(Inf(3)) Then Return "创建人员数组错误：日期"
        If Not isPhone(Inf(4)) Then Return "创建人员数组错误：联系方式"
        If Creater = "" Then Return "错误：创建对象方式不正确"
        Return "ok"
    End Function
    ''' <summary>
    '''  功能说明：卡控数据变更信息
    ''' </summary>
    ''' <returns></returns>
    Private Function changeJudge() As String
        If Creater = "" Then Return "错误：创建对象方式不正确"
        If SearchMsg("Nm") = "" Then Return "错误:该工号不存在人员"
        Return "ok"
    End Function
    ''' <summary>
    ''' 功能说明：依工号查询人员数据表中单列信息
    ''' </summary>
    ''' <param name="SeachType">查询列</param>
    ''' <returns></returns>
    Private Function SearchMsg(ByRef SeachType As String) As String
        Sqlcmd.CommandText = "SELECT " & SeachType & " FROM " & DataBase & " Where No='" & No & "'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    If Not dr(SeachType) Is DBNull.Value Then
                        SearchMsg = dr(SeachType)
                        dr.Close()
                        SqlConn.Close()
                        Return SearchMsg
                    Else
                        dr.Close()
                        SqlConn.Close()
                        Return ""
                    End If
                Else
                    dr.Close()
                    SqlConn.Close()
                    Return ""
                End If
            Else
                dr.Close()
                SqlConn.Close()
                Return ""
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
            Return ex.Message
        End Try
    End Function
    ''' <summary>
    ''' 功能说明：依姓名查询人员数据表中单列信息
    ''' </summary>
    ''' <param name="SeachType">查询列</param>
    ''' <param name="Name">姓名</param>
    ''' <returns></returns>
    Private Function SearchMsg(ByRef SeachType As String, ByRef Name As String) As String
        Sqlcmd.CommandText = "SELECT " + SeachType + " FROM " & DataBase & " Where Nm='" + Name + "'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    If Not dr(SeachType) Is DBNull.Value Then
                        SearchMsg = dr(SeachType)
                        dr.Close()
                        SqlConn.Close()
                        Return SearchMsg
                    Else
                        dr.Close()
                        SqlConn.Close()
                        Return ""
                    End If
                Else
                    dr.Close()
                    SqlConn.Close()
                    Return ""
                End If
            Else
                dr.Close()
                SqlConn.Close()
                Return ""
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
            Return ex.Message
        End Try
    End Function
    ''' <summary>
    ''' 功能说明：依工号查询人员数据表中单列信息
    ''' </summary>
    ''' <param name="SeachType">查询列</param>
    ''' <returns></returns>
    Private Function searchPurview(ByRef SeachType As String, ByRef Duty1 As String) As String
        Sqlcmd.CommandText = "SELECT " & SeachType & " FROM NB_DropDownList Where DropDownListClass='Purview' and DropDownListTpye='" & Department & "' and DropDownListText='" & Duty1 & "'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    If Not dr(SeachType) Is DBNull.Value Then
                        searchPurview = dr(SeachType)
                        dr.Close()
                        SqlConn.Close()
                        Return searchPurview
                    Else
                        dr.Close()
                        SqlConn.Close()
                        Return ""
                    End If
                Else
                    dr.Close()
                    SqlConn.Close()
                    Return ""
                End If
            Else
                dr.Close()
                SqlConn.Close()
                Return ""
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
            Return ex.Message
        End Try
    End Function
    ''' <summary>
    ''' 功能说明：创建新成员
    ''' </summary>
    ''' <returns></returns>
    Function creatEmployee() As String
        If SearchMsg("Nm") = "" Then
            If creatJudge() <> "ok" Then Return creatJudge()
            SqlStr = ""
            ''创建数据
            SqlStr = SqlStr & "Insert into NB_NB410hr_data (Cumulative_attendance,Station,Duty,Grade,Class_1"
            SqlStr = SqlStr & " ,Department_Type,No,Nm,Entry_date,Phone_number) "
            SqlStr = SqlStr & " Values ('0','不分组','作业员','0','-'"
            SqlStr = SqlStr & " ,'" & Inf(0) & "','" & Inf(1) & "','" & Inf(2) & "','" & Inf(3) & "','" & Inf(4) & "')"
            SqlStr = SqlStr & Chr(13)
            ''创建异动数据
            SqlStr = SqlStr & "Insert into NB_Change_Io ( Status_hr,Station_old, Station_new"
            SqlStr = SqlStr & " ,Department_Type,No, Nm, Entry_date, Phone_number"
            SqlStr = SqlStr & " , Transaction_time,Transaction_person) "
            SqlStr = SqlStr & " Values ('人员新增','New_add','不分组'"
            SqlStr = SqlStr & " ,'" & Inf(0) & "','" & Inf(1) & "','" & Inf(2) & "','" & Inf(3) & "','" & Inf(4) & "'"
            SqlStr = SqlStr & " ,'" & Now & "','" & Creater & "')"
            Return ToSql(SqlStr, 2)
        Else
            Return "查询该工号对应为： " & Name & " 请确认！"
        End If
    End Function
    ''' <summary>
    ''' 功能说明：部门变更
    ''' </summary>
    ''' <param name="newStr">新部门</param>
    ''' <returns></returns>
    Function departmentChange(ByRef newStr As String) As String
        If changeJudge() <> "ok" Then Return changeJudge()
        SqlStr = ""
        SqlStr = SqlStr & " Insert into NB_Change_Io ( Department_Type,No, Nm,Status_hr,Dt_old,Dt_new,Station_old,Station_new,Class_old,Class_new"
        SqlStr = SqlStr & " , Transaction_time,Transaction_person) "
        SqlStr = SqlStr & " Select  Department_Type,No,Nm,'部门异动',Department_Type,'" & newStr & "',Station,'不分组',Class_1,'-'"
        SqlStr = SqlStr & " ,'" & Now & "','" & Creater & "' From NB_NB410hr_data"
        SqlStr = SqlStr & "  Where No='" & No & "'"
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & " UPDATE NB_NB410hr_data Set "
        SqlStr = SqlStr & " Department_Type ='" & newStr & "',Station='不分组',Class_1='-',Duty='作业员', Pre_Workspace='-', Pre_Task='-', Workspace='-', Task='-' "
        SqlStr = SqlStr & " Where No='" & No & "'"
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & " DELETE FROM NB_USERS WHERE NB_User_No='" & No & "'"
        Return ToSql(SqlStr, 3)
    End Function
    ''' <summary>
    ''' 功能说明：站别变更
    ''' </summary>
    ''' <param name="newStr">新站别</param>
    ''' <returns></returns>
    Function stationChange(ByRef newStr As String) As String
        If changeJudge() <> "ok" Then Return changeJudge()
        SqlStr = ""
        SqlStr = SqlStr & " Insert into NB_Change_Io ( Department_Type,No, Nm,Status_hr, Station_old, Station_new,Class_old"
        SqlStr = SqlStr & " , Transaction_time,Transaction_person) "
        SqlStr = SqlStr & " Select  Department_Type,No,Nm,'站别异动',Station,'" & newStr & "',Class_1"
        SqlStr = SqlStr & " ,'" & Now & "','" & Creater & "' From NB_NB410hr_data"
        SqlStr = SqlStr & "  Where No='" & No & "'"
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & " UPDATE NB_NB410hr_data Set "
        SqlStr = SqlStr & " Station ='" & newStr & "',Duty='作业员', Pre_Workspace='-', Pre_Task='-', Workspace='-', Task='-' "
        SqlStr = SqlStr & " Where No='" & No & "'"
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & " DELETE FROM NB_USERS WHERE NB_User_No='" & No & "'"
        Return ToSql(SqlStr, 3)
    End Function
    ''' <summary>
    ''' 功能说明：班别变更
    ''' </summary>
    ''' <param name="newStr">新班别</param>
    ''' <returns></returns>
    Function classChange(ByRef newStr As String) As String
        If changeJudge() <> "ok" Then Return changeJudge()
        SqlStr = ""
        SqlStr = SqlStr & " Insert into NB_Change_Io ( Department_Type,No,Nm,Status_hr, Station_old,Class_old,Class_new"
        SqlStr = SqlStr & " , Transaction_time,Transaction_person) "
        SqlStr = SqlStr & " Select  Department_Type,No,Nm,'班别异动',Station,Class_1,'" & newStr & "'"
        SqlStr = SqlStr & " ,'" & Now & "','" & Creater & "' From NB_NB410hr_data"
        SqlStr = SqlStr & "  Where No='" & No & "'"
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & " UPDATE NB_NB410hr_data Set "
        SqlStr = SqlStr & " Class_1 ='" & newStr & "',Duty='作业员', Pre_Workspace='-', Pre_Task='-', Workspace='-', Task='-' "
        SqlStr = SqlStr & " Where No='" & No & "'"
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & " DELETE FROM NB_USERS WHERE NB_User_No='" & No & "'"
        Return ToSql(SqlStr, 3)
    End Function
    ''' <summary>
    ''' 功能说明：职务变更
    ''' </summary>
    ''' <param name="newStr">新职务</param>
    ''' <returns></returns>
    Function positionChange(ByRef newStr As String) As String
        If changeJudge() <> "ok" Then Return changeJudge()
        Dim i As Byte = 0
        Dim newPurview As String = searchPurview("AddUser", newStr)
        SqlStr = ""
        If newPurview = "" Then
            SqlStr = SqlStr & " DELETE FROM NB_USERS WHERE NB_User_No='" & No & "'"
            i = 1
        ElseIf Purview = "" And newPurview <> "" Then
            SqlStr = SqlStr & " Insert into NB_USERS (NB_User_No, NB_User_Name, NB_User_PWD, NB_User_Purview, NB_User_Group) values ('" & No & "','" & Name & "','" & No & "','" & newStr & "','" & newPurview & "')"
            i = 1
        ElseIf Purview <> newPurview Then
            SqlStr = SqlStr & " UPDATE NB_USERS Set NB_User_Purview ='" & newStr & "',NB_User_Group='" & newPurview & "' where NB_User_No='" & No & "'"
            i = 1
        End If
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & " Insert into NB_Change_Io ( Department_Type,No,Nm,Status_hr,Station_old,Class_old,Duty_old,Duty_new, Purview_new"
        SqlStr = SqlStr & " , Transaction_time,Transaction_person) "
        SqlStr = SqlStr & " Select  Department_Type,No,Nm,'职务调整',Station,Class_1,Duty,'" & newStr & "','" & newPurview & "'"
        SqlStr = SqlStr & " ,'" & Now & "','" & Creater & "' From NB_NB410hr_data"
        SqlStr = SqlStr & "  Where No='" & No & "'"
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & " UPDATE NB_NB410hr_data Set "
        SqlStr = SqlStr & " Duty='" & newStr & "', Pre_Workspace='-', Pre_Task='-', Workspace='-', Task='-' "
        SqlStr = SqlStr & " Where No='" & No & "'"
        SqlStr = SqlStr & Chr(13)
        i = i + 2
        Return ToSql(SqlStr, i)
    End Function
    ''' <summary>
    ''' 功能说明：提交Sql数据库
    ''' </summary>
    ''' <param name="SqlStr">sql语句</param>
    ''' <param name="Item">提交数量</param>
    ''' <returns></returns>
    Protected Function ToSql(ByRef SqlStr As String, ByRef Item As Byte) As String
        Dim i As Byte = 0
        SqlConn.Open()
        Trans = SqlConn.BeginTransaction '第一步，开始事务。这句必须在下句前，相当于实例事务对象
        Sqlcmd.Transaction = Trans     'Sqlcmd要做的事
        Try
            Sqlcmd.CommandText = SqlStr
            i = Sqlcmd.ExecuteNonQuery()
            Trans.Commit()
            SqlConn.Close()
            ToSql = "提交成功! " & i & " / " & Item '& SqlStr
        Catch ex As Exception
            Trans.Rollback()
            SqlConn.Close()
            ToSql = "提交失败: " & ex.Message.ToString & vbCrLf & SqlStr
        End Try
    End Function
End Class
