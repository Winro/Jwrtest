Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient


Public Class Cutters
    Dim SqlConn As New SqlConnection(ConfigurationManager.ConnectionStrings("NB410Sys").ConnectionString)
    Dim Sqlcmd As SqlCommand = SqlConn.CreateCommand
    Dim Trans As SqlTransaction
    Dim dr As SqlDataReader
    Dim SqlStr As String = ""
    Private Name As String
    Private UseOperator As String
    Private SizeOrder As Byte
    Public HasCutterStatus As Boolean = False
    Public Sub New()
    End Sub
    ''' <summary>
    ''' 功能说明：构造刀模对象
    ''' </summary>
    ''' <param name="CutterName">刀模名称</param>
    Public Sub New(ByRef CutterName As String)
        Name = UCase(CutterName.Trim)
        HasCutterStatus = HasCutter()
    End Sub
    ''' <summary>
    ''' 功能说明：构造刀模对象
    ''' </summary>
    ''' <param name="CutterName">刀模名称</param>
    ''' <param name="User">构造人员</param>
    Public Sub New(ByRef CutterName As String, ByRef User As String)
        Name = UCase(CutterName.Trim)
        UseOperator = User
        HasCutterStatus = HasCutter()
    End Sub
    ''' <summary>
    ''' 功能说明：构造刀模对象
    ''' </summary>
    ''' <param name="CutterName">刀模名称</param>
    ''' <param name="User">构造人员</param>
    ''' <param name="Order">尺寸序号</param>
    Public Sub New(ByRef CutterName As String, ByRef User As String, ByRef Order As Byte)
        Name = UCase(CutterName.Trim)
        UseOperator = User
        SizeOrder = Order
        HasCutterStatus = HasCutter()
    End Sub
    ''' <summary>
    ''' 属性说明：刀模使用次数
    ''' </summary>
    ''' <returns></returns>
    Property Use_Times As Integer
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Use_times")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Integer)
            UpdateMsg_Basic("Dm_Use_times", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模维修次数
    ''' </summary>
    ''' <returns></returns>
    Property Maintenance_Times As Integer
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Maintenance_times")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Integer)
            UpdateMsg_Basic("Dm_Maintenance_times", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：中继片对应刀模号
    ''' </summary>
    ''' <returns></returns>
    Property Oppsite_CutterName As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Oppsite")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_Oppsite", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：入厂日期
    ''' </summary>
    ''' <returns></returns>
    Property Entry_Date As Date
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Date_of_entry")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Date)
            UpdateMsg_Basic("Dm_Date_of_entry", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模凹槽\导角\耳朵状态
    ''' </summary>
    ''' <returns></returns>
    Property Groove As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Groove")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Groove", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模使用状态
    ''' </summary>
    ''' <returns></returns>
    Property Using_Status As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_status")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_status", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模储位区域位置
    ''' </summary>
    ''' <returns></returns>
    Property Location As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Status_1")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_Status_1", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模储位区域位置具体
    ''' </summary>
    ''' <returns></returns>
    Property Location_Point As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Status_2")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_Status_2", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：机台限制
    ''' </summary>
    ''' <returns></returns>
    Property Using_Limit As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("RTC_Limit_use_machine")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("RTC_Limit_use_machine", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模类型
    ''' </summary>
    ''' <returns></returns>
    Property Type As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_type")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_type", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模尺寸组成
    ''' </summary>
    ''' <returns></returns>
    Property Size_Form As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Size_all")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_Size_all", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模刀板尺寸
    ''' </summary>
    ''' <returns></returns>
    Property Size As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Cutter_size")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_Cutter_size", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模有效幅宽
    ''' </summary>
    ''' <returns></returns>
    Property Breadth As Double
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_breadth")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Double)
            UpdateMsg_Basic("Dm_breadth", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模设计幅宽
    ''' </summary>
    ''' <returns></returns>
    Property Design_Breadth As Double
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_design_breadth")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Double)
            UpdateMsg_Basic("Dm_design_breadth", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模有效移送长
    ''' </summary>
    ''' <returns></returns>
    Property Effective_Length As Double
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Effective_length")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Double)
            UpdateMsg_Basic("Dm_Effective_length", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模设计喷印总数量
    ''' </summary>
    ''' <returns></returns>
    Property Print_Quantity As Integer
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Jet_printing_shaft_number")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Integer)
            UpdateMsg_Basic("Dm_Jet_printing_shaft_number", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模设计蓝喷印数量
    ''' </summary>
    ''' <returns></returns>
    Property Print_Quantity_Blue As Integer
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Jet_printing_shaft_number_blue")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Integer)
            UpdateMsg_Basic("Dm_Jet_printing_shaft_number_blue", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模设计红喷印数量
    ''' </summary>
    ''' <returns></returns>
    Property Print_Quantity_Red As Integer
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_Jet_printing_shaft_number_red")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Integer)
            UpdateMsg_Basic("Dm_Jet_printing_shaft_number_red", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：刀模备注
    ''' </summary>
    ''' <returns></returns>
    Property Remarks As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("Dm_remarks")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_remarks", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：RTC裁切刀模备注
    ''' </summary>
    ''' <returns></returns>
    Property RTC_Notice As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("RTC_Notice")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("RTC_Notice", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明：RGR裁切刀模备注
    ''' </summary>
    ''' <returns></returns>
    Property RGR_Notice As String
        Get
            If HasCutter() Then
                Return SearchMsg_Basic("RGR_Notice")
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("RGR_Notice", value)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明:尺寸-尺寸
    ''' </summary>
    ''' <returns></returns>
    Property Size_Str As String
        Get
            If HasCutter() Then
                If SizeOrder = "" Then
                    Return SearchMsg_Basic("Dm_size", SizeOrder)
                Else
                    Return "请输入尺寸序列编号"
                End If
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_size", value, SizeOrder)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明:尺寸-启用状态
    ''' </summary>
    ''' <returns></returns>
    Property Size_Status As String
        Get
            If HasCutter() Then
                If SizeOrder = "" Then
                    Return SearchMsg_Basic("Dm_size_status", SizeOrder)
                Else
                    Return "请输入尺寸序列编号"
                End If
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As String)
            UpdateMsg_Basic("Dm_size_status", value, SizeOrder)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明:尺寸-取片数
    ''' </summary>
    ''' <returns></returns>
    Property Size_Quantity As Integer
        Get
            If HasCutter() Then
                If SizeOrder = "" Then
                    Return SearchMsg_Basic("Dm_size_quantity", SizeOrder)
                Else
                    Return "请输入尺寸序列编号"
                End If
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Integer)
            UpdateMsg_Basic("Dm_size_quantity", value, SizeOrder)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明:尺寸-刀模尺寸长
    ''' </summary>
    ''' <returns></returns>
    Property Size_Length As Double
        Get
            If HasCutter() Then
                If SizeOrder = "" Then
                    Return SearchMsg_Basic("Dm_size_length", SizeOrder)
                Else
                    Return "请输入尺寸序列编号"
                End If
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Double)
            UpdateMsg_Basic("Dm_size_length", value, SizeOrder)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明:尺寸-刀模尺寸宽
    ''' </summary>
    ''' <returns></returns>
    Property Size_Wide As Double
        Get
            If HasCutter() Then
                If SizeOrder = "" Then
                    Return SearchMsg_Basic("Dm_size_wide", SizeOrder)
                Else
                    Return "请输入尺寸序列编号"
                End If
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Double)
            UpdateMsg_Basic("Dm_size_wide", value, SizeOrder)
        End Set
    End Property
    ''' <summary>
    ''' 属性说明:尺寸-刀模裁切角度
    ''' </summary>
    ''' <returns></returns>
    Property Size_Angle As Double
        Get
            If HasCutter() Then
                If SizeOrder = "" Then
                    Return SearchMsg_Basic("Dm_size_Absorption_axis", SizeOrder)
                Else
                    Return "请输入尺寸序列编号"
                End If
            Else
                Return "查无" & Name & "刀模"
            End If
        End Get
        Set(value As Double)
            UpdateMsg_Basic("Dm_size_Absorption_axis", value, SizeOrder)
        End Set
    End Property
    ''' <summary>
    ''' 功能说明：判定刀模是否存在
    ''' </summary>
    ''' <returns></returns>
    Function HasCutter() As Boolean
        If SearchMsg_Basic("Dm_number").ToString Like "?#####-##" Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 功能说明：查询NB_Dm_Basic表中数据
    ''' </summary>
    ''' <param name="SeachType">查询项目</param>
    ''' <returns></returns>
    Private Function SearchMsg_Basic(ByRef SeachType As String) As String
        If Format() Then
            Sqlcmd.CommandText = "SELECT " + SeachType + " FROM NB_Dm_Basic Where Dm_number='" + Name + "'"
            SearchMsg_Basic = ""
            SqlConn.Open()
            Try
                dr = Sqlcmd.ExecuteReader
                If dr.HasRows Then
                    If dr.Read Then
                        If Not dr(SeachType) Is DBNull.Value Then
                            SearchMsg_Basic = dr(SeachType).ToString
                            dr.Close()
                            SqlConn.Close()
                            Return SearchMsg_Basic
                        End If
                    End If
                End If
                dr.Close()
                SqlConn.Close()
            Catch ex As Exception
                If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
                SqlConn.Close()
                Return "错误：" & ex.Message.ToString
            End Try
        Else
            Return Format()
        End If
    End Function
    ''' <summary>
    ''' 功能说明：查询NB_Dm_Basic_Size表中数据
    ''' </summary>
    ''' <param name="SeachType">查询项目</param>
    ''' <returns></returns>
    Private Function SearchMsg_Basic(ByRef SeachType As String, ByRef Order As Byte) As String
        If Format() Then
            Sqlcmd.CommandText = "SELECT " + SeachType + " FROM NB_Dm_Basic_Size Where Dm_number='" + Name + "' and Dm_size_order='" & Order & "'"
            SearchMsg_Basic = ""
            SqlConn.Open()
            Try
                dr = Sqlcmd.ExecuteReader
                If dr.HasRows Then
                    If dr.Read Then
                        If Not dr(SeachType) Is DBNull.Value Then
                            SearchMsg_Basic = dr(SeachType).ToString
                            dr.Close()
                            SqlConn.Close()
                            Return SearchMsg_Basic
                        End If
                    End If
                End If
                dr.Close()
                SqlConn.Close()
            Catch ex As Exception
                If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
                SqlConn.Close()
                Return "错误：" & ex.Message.ToString
            End Try
        Else
            Return Format()
        End If
    End Function
    ''' <summary>
    ''' 功能说明：更新NB_Dm_Basic表中数据
    ''' </summary>
    ''' <param name="Item">项目</param>
    ''' <param name="Value">值</param>
    ''' <returns></returns>
    Private Function UpdateMsg_Basic(ByRef Item As String, ByRef Value As String) As String
        If UseOperator = "" Then Return "错误：操作人员不许为空"
        Dim SqlStr As String = ""
        ''插入流水记录-新增属性变更记录
        SqlStr = SqlStr & "insert into NB_DM_io (Dm_number, Dm_Transaction_time, Dm_Transaction_person"
        SqlStr = SqlStr & " , Status, Dm_Reservoir_type_before, Dm_Reservoir_number_before"
        SqlStr = SqlStr & " , Dm_Reservoir_type_after, Dm_Reservoir_number_after) "
        SqlStr = SqlStr & " Select Dm_number,'" + Now + "','" + UseOperator + "'"
        SqlStr = SqlStr & " ,'基本属性变更','" & Item & "'," & Item & ""
        SqlStr = SqlStr & " ,'" & Item & "','" & Value & "'"
        SqlStr = SqlStr & " From NB_Dm_Basic Where Dm_number='" + Name + "'"
        ''更新库存表-物品属性
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & "update NB_Dm_Basic Set " & Item & "= '" & Value & "' Where Dm_number='" + Name + "'"
        Return ToSql(SqlStr, 2)
    End Function
    ''' <summary>
    ''' 功能说明：更新NB_Dm_Basic_Size表中数据
    ''' </summary>
    ''' <param name="Item">项目</param>
    ''' <param name="Value">值</param>
    ''' <param name="Order">尺寸编号</param>
    ''' <returns></returns>
    Private Function UpdateMsg_Basic(ByRef Item As String, ByRef Value As String, ByRef Order As Byte) As String
        If UseOperator = "" Then Return "错误：操作人员不许为空"
        If Order = "" Then Return "错误：尺寸序号不许为空"
        Dim SqlStr As String = ""
        ''插入流水记录-新增属性变更记录
        SqlStr = SqlStr & "insert into NB_DM_io (Dm_number, Dm_Transaction_time, Dm_Transaction_person"
        SqlStr = SqlStr & " , Status, Dm_Reservoir_type_before, Dm_Reservoir_number_before"
        SqlStr = SqlStr & " , Dm_Reservoir_type_after, Dm_Reservoir_number_after,Dm_Remark) "
        SqlStr = SqlStr & " Select Dm_number,'" + Now + "','" + UseOperator + "'"
        SqlStr = SqlStr & " ,'尺寸属性变更','" & Item & "'," & Item & ""
        SqlStr = SqlStr & " ,'" & Item & "','" & Value & "','" & Order & "'"
        SqlStr = SqlStr & " From NB_Dm_Basic_Size Where Dm_number='" + Name + "' and Dm_size_order='" & Order & "'"
        ''更新库存表-物品属性
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & "update NB_Dm_Basic_Size Set " & Item & "= '" & Value & "' Where Dm_number='" + Name + "' and Dm_size_order='" & Order & "'"
        Return ToSql(SqlStr, 2)
    End Function
    ''' <summary>
    ''' 功能说明：上传至数据库
    ''' </summary>
    ''' <param name="SqlStr">SQL语句</param>
    ''' <param name="Item">上传语句数量</param>
    ''' <returns></returns>
    Private Function ToSql(ByRef SqlStr As String, ByRef Item As Byte) As String
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
    ''' <summary>
    ''' 功能说明：获取所有状态尺寸表
    ''' </summary>
    ''' <returns></returns>
    Function GetSizeTable() As DataTable
        SqlStr = ""
        SqlStr = SqlStr & "SELECT ID, Dm_number , Dm_size "
        SqlStr = SqlStr & "  , Dm_size_status , Dm_size_quantity "
        SqlStr = SqlStr & "  , Dm_size_length , Dm_size_wide"
        SqlStr = SqlStr & "  , Dm_size_Absorption_axis , Dm_size_order "
        SqlStr = SqlStr & "   FROM NB_Dm_Basic_Size Where Dm_number='" + Name + "'  order by Dm_size_order"
        Return CreatTable("Tb_Size", SqlStr)
    End Function
    ''' <summary>
    ''' 功能说明：获取特定序列编号尺寸表
    ''' </summary>
    ''' <returns></returns>
    Function GetSizeTable(ByRef Order As Byte) As DataTable
        SqlStr = ""
        SqlStr = SqlStr & "SELECT ID, Dm_number , Dm_size "
        SqlStr = SqlStr & "  , Dm_size_status , Dm_size_quantity "
        SqlStr = SqlStr & "  , Dm_size_length , Dm_size_wide"
        SqlStr = SqlStr & "  , Dm_size_Absorption_axis , Dm_size_order "
        SqlStr = SqlStr & "   FROM NB_Dm_Basic_Size Where Dm_number='" + Name + "'  and Dm_size_order='" & Order & "'  order by Dm_size_order"
        Return CreatTable("Tb_Size", SqlStr)
    End Function
    ''' <summary>
    ''' 功能说明：获取启用状态尺寸表
    ''' </summary>
    ''' <returns></returns>
    Function GetSizeTable(ByRef SizeStatus As String) As DataTable
        SqlStr = ""
        SqlStr = SqlStr & "SELECT ID, Dm_number , Dm_size "
        SqlStr = SqlStr & "  , Dm_size_status , Dm_size_quantity "
        SqlStr = SqlStr & "  , Dm_size_length , Dm_size_wide"
        SqlStr = SqlStr & "  , Dm_size_Absorption_axis , Dm_size_order "
        SqlStr = SqlStr & "   FROM NB_Dm_Basic_Size Where Dm_number='" + Name + "'  and Dm_size_status = '" & SizeStatus & "'  order by Dm_size_order"
        Return CreatTable("Tb_Size", SqlStr)
    End Function
    ''' <summary>
    ''' 功能说明：获取所有状态尺寸表,中文标题
    ''' </summary>
    ''' <returns></returns>
    Function GetSizeTable(ByRef Chinese As Boolean) As DataTable
        SqlStr = ""
        SqlStr = SqlStr & "SELECT ID, Dm_number , Dm_size "
        SqlStr = SqlStr & "  , Dm_size_status , Dm_size_quantity "
        SqlStr = SqlStr & "  , Dm_size_length , Dm_size_wide"
        SqlStr = SqlStr & "  , Dm_size_Absorption_axis , Dm_size_order "
        SqlStr = SqlStr & "   FROM NB_Dm_Basic_Size Where Dm_number='" + Name + "'  order by Dm_size_order"
        Return CreatTable("Tb_Size", SqlStr, Chinese)
    End Function
    ''' <summary>
    ''' 功能说明：获取特定序列编号尺寸表,中文标题
    ''' </summary>
    ''' <returns></returns>
    Function GetSizeTable(ByRef Order As Byte, ByRef Chinese As Boolean) As DataTable
        SqlStr = ""
        SqlStr = SqlStr & "SELECT ID, Dm_number , Dm_size "
        SqlStr = SqlStr & "  , Dm_size_status , Dm_size_quantity "
        SqlStr = SqlStr & "  , Dm_size_length , Dm_size_wide"
        SqlStr = SqlStr & "  , Dm_size_Absorption_axis , Dm_size_order "
        SqlStr = SqlStr & "   FROM NB_Dm_Basic_Size Where Dm_number='" + Name + "'  and Dm_size_order='" & Order & "'  order by Dm_size_order"
        Return CreatTable("Tb_Size", SqlStr, Chinese)
    End Function
    ''' <summary>
    ''' 功能说明：获取启用状态尺寸表,中文标题
    ''' </summary>
    ''' <returns></returns>
    Function GetSizeTable(ByRef SizeStatus As String, ByRef Chinese As Boolean) As DataTable
        SqlStr = ""
        SqlStr = SqlStr & "SELECT ID, Dm_number , Dm_size "
        SqlStr = SqlStr & "  , Dm_size_status , Dm_size_quantity "
        SqlStr = SqlStr & "  , Dm_size_length , Dm_size_wide"
        SqlStr = SqlStr & "  , Dm_size_Absorption_axis , Dm_size_order "
        SqlStr = SqlStr & "   FROM NB_Dm_Basic_Size Where Dm_number='" + Name + "'  and Dm_size_status = '" & SizeStatus & "'  order by Dm_size_order"
        Return CreatTable("Tb_Size", SqlStr, Chinese)
    End Function
    ''' <summary>
    '''  功能说明：创建尺寸表
    ''' </summary>
    ''' <param name="SqlStr">Sql语句</param>
    ''' <returns></returns>
    Private Function CreatTable(ByRef TbName As String, ByRef SqlStr As String) As DataTable
        Dim da As New SqlDataAdapter(SqlStr, SqlConn)
        Dim ds As New DataSet
        da.Fill(ds, TbName)
        CreatTable = ds.Tables(TbName)
        Return CreatTable
    End Function
    ''' <summary>
    '''  功能说明：创建尺寸表，中文标题
    ''' </summary>
    ''' <param name="SqlStr">Sql语句</param>
    ''' <returns></returns>
    Private Function CreatTable(ByRef TbName As String, ByRef SqlStr As String, ByRef Chinese As Boolean) As DataTable
        Dim da As New SqlDataAdapter(SqlStr, SqlConn)
        Dim ds As New DataSet
        da.Fill(ds, TbName)
        CreatTable = ds.Tables(TbName)
        If Chinese Then
            CreatTable = ChiTransformEng(CreatTable)
        End If
        Return CreatTable
    End Function
    ''' <summary>
    ''' 功能说明：判定输入刀模格式是否正确
    ''' </summary>
    ''' <returns></returns>
    Private Function Format() As Boolean
        If Name Like "?#####-##" Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 数据表中英对照，替换标题为中文
    ''' </summary>
    ''' <param name="Dt">数据表</param>
    ''' <returns></returns>
    Private Function ChiTransformEng(ByRef Dt As DataTable) As DataTable
        Dim arr() As String = {}
        Select Case Dt.TableName
            Case "Tb_Size"
                arr = {"ID|序号", "Dm_number|刀模流水号", "Dm_size|尺寸", "Dm_size_status|使用状态", "Dm_size_quantity|取片数", "Dm_size_length|尺寸长", "Dm_size_wide|尺寸宽", "Dm_size_Absorption_axis|裁切角度", "Dm_size_order|尺寸序号"}
            Case "Tb_Basic"
                arr = {"ID|序号", "Dm_number|刀模流水号", "Dm_Oppsite|对应刀模", "Dm_Date_of_entry|创建日期", "Dm_Use_times|使用次数", "Dm_Maintenance_times|维修次数", "Groove|凹槽状况", "Dm_status|启用状况", "Dm_Status_1|储区", "Dm_Status_2|具体储位", "Dm_Status_3|无", "Space_Number|无", "Space_Use_type|无", "RTC_Limit_use_machine|限制机台", "Dm_type|类型", "Dm_Size_all|尺寸组合", "Dm_Cutter_size|刀板尺寸", "Dm_breadth|有效幅宽", "Dm_design_breadth|设计幅宽", "Dm_Effective_length|有效移送长", "Dm_Jet_printing_shaft_number|喷印机数量", "Dm_Jet_printing_shaft_number_blue|喷印机数量(蓝)", "Dm_Jet_printing_shaft_number_red|喷印机数量(红)", "Dm_remarks|备注", "RTC_Notice|裁切注意事项", "RGR_Notice|磨边注意事项"}
            Case "Tb_Inspection"
                arr = {"ID|序号", "Dm_Number|刀模流水号", "Use_Times|使用次数", "Inspection_Finish_Status|无", "Inspection_Status|无", "Inspection_Type|检验类型", "Dm_Maintenance_times|检验次数", "Inspection_Project|检验项目", "Inspection_Finish|无", "CellPoint|检验点位", "Inspection_Size|检验尺寸", "Test_Results_Int|量测值", "Test_Results_String|检验结果", "Standard_Max|标准上限", "Standard_Min|标准下限", "Standard_String|标准", "Standard_Length|标准长", "Standard_Wide|标准宽", "Standard_Squareness|标准直角度", "Tolerance_Length|公差长", "Tolerance_Wide|公差宽", "Tolerance_Squareness|公差直角度", "Measured_Length_1|量测长1", "Measured_Wide_1|量测宽1", "Measured_Length_2|量测长2", "Measured_Wide_2|量测宽2", "Measured_Length_3|量测长3", "Measured_Wide_3|量测宽3", "Measured_Squareness_1|直角度1", "Measured_Squareness_2|直角度2", "Measured_Squareness_3|直角度3", "Measured_Squareness_4|直角度4", "Inspection_personnel|检验人", "Inspection_date|检验时间", "remarks|备注", "CreatPerson|创建人", "CreatTime|创建时间", "inspection_and_audit_staff|审核人", "inspection_and_audit_staff_date|审核时间", "Transaction_time|异动时间", "Transaction_person|异动人"}
            Case "Tb_IO"
                arr = {"Id|序号", "Dm_number|刀模流水号", "Dm_Transaction_time|变更时间", "Dm_Transaction_person|变更人", "Status|状态", "Dm_Reservoir_type_before|变更类型（变更前）", "Dm_Reservoir_number_before|变更值（变更前）", "Dm_Reservoir_type_after|变更类型（变更后）", "Dm_Reservoir_number_after|变更值（变更后）", "Dm_Preparation_tool|工具", "Dm_Remark|备注"}
        End Select
        For m As Integer = 0 To Dt.Columns.Count - 1
            For n As Integer = LBound(arr) To UBound(arr)
                If Dt.Columns(m).Caption = Split(arr(n), "|")(0).ToString Then
                    Dt.Columns(m).ColumnName = Split(arr(n), "|")(1).ToString
                    Exit For
                End If
            Next
        Next
        Return Dt
    End Function
    Enum err
        查无刀模 = 0
        ValueTwo
    End Enum
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
