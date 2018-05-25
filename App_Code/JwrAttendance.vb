Imports Microsoft.VisualBasic

Public Class JwrAttendance
    Inherits JwrEmployee
    Private SqlStr As String = ""
    Private Nt As Date = Ntd()
    Public Sub New()
    End Sub
    ''' <summary>
    ''' 功能说明：创建对象
    ''' </summary>
    ''' <param name="department_c">部门</param>
    '''  <param name="station_c">站别</param>
    '''  <param name="Class_1_c">班别</param>
    ''' <param name="Creater_c">创建人</param>
    Public Sub New(ByVal department_c As String, ByRef station_c As String, ByRef class_1_c As String, ByRef creater_c As String)
        Department = department_c
        Station = station_c
        Class_1 = class_1_c
        Creater = creater_c
    End Sub
    ''' <summary>
    ''' 功能说明：提供当日作业日期
    ''' </summary>
    ''' <returns></returns>
    Public Function Ntd() As Date
        Dim Nnt As Date = Now
        Dim Nt As Date '点名日期
        Dim NL As Int64
        NL = DateDiff("s", Format(Nnt, "yyyy/MM/dd" & " 07:50:00"), Nnt)
        If NL < 0 Then
            Nt = CDate(Format(Nnt.AddDays(-1), "yyyy/MM/dd"))
        ElseIf NL >= 0 Then
            Nt = CDate(Format(Nnt, "yyyy/MM/dd"))
        End If
        'If Me.RadioButtonList_Class.SelectedValue = "B班" Then Nt = "2017/04/27"
        'Nt = "2017/03/16"
        Return Nt
    End Function
    Private Function judgeCreatAttendaceSheet(ByRef aDate As Date) As String
        If Not isDepartment(Department) Then Return "创建错误：部门格式异常"
        If Not IsDate(aDate) Then Return "创建错误：日期格式异常"
        If Len(Station) = 0 Then Return "创建错误：站点为空"
        If Len(Class_1) = 0 Then Return "创建错误：班别为空"
        If Len(Creater) = 0 Then Return "创建错误：创建者为空"
        Return "ok"
    End Function
    Function creatAttendaceSheet(ByRef aDate As Date) As String
        If judgeCreatAttendaceSheet(aDate) <> "ok" Then Return judgeCreatAttendaceSheet(aDate)
        Dim i As Byte = 0
        SqlStr = ""
        If aDate = Nt Then
            SqlStr = SqlStr & " Insert into NB_Hr_Attendance_io (Department_Type, No, Nm, Station, Class_1, Call_Date_1, Attendance_Date, State_of_Attendance, Transaction_time, Transaction_person) "
            SqlStr = SqlStr & " Select '" & Department & "',No, Nm, Station, Class_1 ,'" & Nt & "', '" + Now + "','-', '" + Now + "','" & Creater & "'"
            SqlStr = SqlStr & " From NB_NB410hr_data "
            SqlStr = SqlStr & " Where Station='" + Station + "' and Class_1='" + Class_1 + "'  and (State_of_Attendance is NULL or State_of_Attendance <> '离职') and no not in ( Select No from NB_Hr_Attendance_io where Call_Date_1= '" + Nt + "')"
            SqlStr = SqlStr & Chr(13)
            i = 1
        End If
        SqlStr = SqlStr & " Update NB_NB410hr_data set Temp_Cumulative_attendance='0' "
        SqlStr = SqlStr & " where  Call_Date_1<'" + Nt.AddDays(-1) + "' and  Station='" + Station + "'"
        SqlStr = SqlStr & " and Class_1='" + Class_1 + "'  and State_of_Attendance<>'离职'"
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & " Update NB_NB410hr_data set "
        SqlStr = SqlStr & " Cumulative_attendance=(CASE WHEN Temp_Cumulative_attendance=1 THEN (Cumulative_attendance+Temp_Cumulative_attendance) ELSE '0' END)"
        SqlStr = SqlStr & " ,Temp_Cumulative_attendance='0',Call_Date_Temp='" + Nt + "'"
        SqlStr = SqlStr & " where  (Call_Date_Temp<'" + Nt + "' OR Call_Date_Temp IS NULL)"
        SqlStr = SqlStr & " and  Station='" + Station + "' and Class_1='" + Class_1 + "'  and State_of_Attendance<>'离职'"
        i = i + 2
        Return ToSql(SqlStr, i)
    End Function
    Function naming(ByRef no1 As String, ByRef newState As String, ByRef backup As String) As String
        Dim c As Byte = 0
        Select Case newState
            Case "出勤", "下早班"
                c = 1
        End Select
        SqlStr = ""
        SqlStr = SqlStr & " Insert into NB_Hr_Attendance_io (Backup_1,Department_Type,No,Nm,Station,Class_1"
        SqlStr = SqlStr & " ,Call_Date_1,  Attendance_Date, State_of_Attendance, Transaction_time, Transaction_person)"
        SqlStr = SqlStr & " Select '" & backup & "',Department_Type,No,Nm,Station,Class_1"
        SqlStr = SqlStr & " "
        'Dim Str_Values, strSQL As String
        'Str_Values = " '" + TextBox_Backup_1.Text + "','" & Session("Dt_hr") & "','" + Gvr.Cells(1).Text + "','" + Gvr.Cells(2).Text + "','" + Station + "','" + Class_1 + "','" & Me.TextBox_Date.Text & "','" + Now + "','" + Rb.Text + "','" + Now + "','" + Session("UserName_hr") + "'"
        'strSQL = "Insert into NB_Hr_Attendance_io (Backup_1,Department_Type,No,Nm,Station,Class_1,Call_Date_1,  Attendance_Date, State_of_Attendance, Transaction_time, Transaction_person) values ( " & Str_Values & " )"
        'With Me.SqlDataSource_hr
        '    .UpdateCommand = "UPDATE NB_NB410hr_data set State_of_Attendance='" + Rb.Text + "',Attendance_date='" + Now + "',Temp_Cumulative_attendance='" & c & "',Call_Date_1='" & Me.TextBox_Date.Text & "' where No='" + Gvr.Cells(1).Text + "'"
        '    .Update()
        '    If Rb.SelectedValue = "离职" Then
        '        .UpdateCommand = "UPDATE NB_NB410hr_data set Pre_Workspace='-', Pre_Task='-', Workspace='-', Task='-' where No='" + Gvr.Cells(1).Text + "'"
        '        .Update()
        '        .DeleteCommand = "DELETE FROM NB_USERS  where NB_User_No='" + Gvr.Cells(1).Text + "'"
        '        .Delete()
        '    End If
        '    .InsertCommand = strSQL
        '    .Insert()
        'End With
    End Function
End Class
