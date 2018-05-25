Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class ConsumptiveMaterial
    Private nm As String = ""
    Dim SqlConn As New SqlConnection(ConfigurationManager.ConnectionStrings("NB410Sys").ConnectionString)
    Dim Sqlcmd As SqlCommand = SqlConn.CreateCommand
    Dim Trans As SqlTransaction
    Dim dr As SqlDataReader
    Dim SqlStr As String = ""
    Private GoodsSN As String
    Private UseOperator As String
    Public Sub New()
    End Sub
    Public Sub New(ByVal SN As String)
        GoodsSN = SN
    End Sub
    Public Sub New(ByVal SN As String, ByRef User As String)
        GoodsSN = SN
        UseOperator = User
    End Sub
    Property GoodsType As String
        Get
            Return SearchGoodsMsg("Nm_Type")
        End Get
        Set(value As String)
            UpdateGoodsMsg("Nm_Type", value)
        End Set
    End Property
    Property GoodsName As String
        Get
            Return SearchGoodsMsg("NM")
        End Get
        Set(value As String)
            UpdateGoodsMsg("NM", value)
        End Set
    End Property
    Property Department_Type As String
        Get
            Return SearchGoodsMsg("Department_Type")
        End Get
        Set(value As String)
            UpdateGoodsMsg("Department_Type", value)
        End Set
    End Property
    Property GoodsState As String
        Get
            Return SearchGoodsMsg("State")
        End Get
        Set(value As String)
            UpdateGoodsMsg("State", value)
        End Set
    End Property
    Property Remarks As String
        Get
            Return SearchGoodsMsg("Remarks")
        End Get
        Set(value As String)
            UpdateGoodsMsg("Remarks", value)
        End Set
    End Property
    Property Quantity_1F As Integer
        Get
            Return GoodsQuantity("Inventory_quantity_1F")
        End Get
        Set(value As Integer)
            UpdateGoodsMsg("Inventory_quantity_1F", value)
        End Set
    End Property
    Property Quantity_Warehouse As Integer
        Get
            Return GoodsQuantity("Warehouse_stock_quantity")
        End Get
        Set(value As Integer)
            UpdateGoodsMsg("Warehouse_stock_quantity", value)
        End Set
    End Property
    Property Purchase_quantity As Integer
        Get
            Return GoodsQuantity("Purchase_quantity")
        End Get
        Set(value As Integer)
            UpdateGoodsMsg("Purchase_quantity", value)
        End Set
    End Property
    Property Safety_stock As Integer
        Get
            Return GoodsQuantity("Safety_stock")
        End Get
        Set(value As Integer)
            UpdateGoodsMsg("Safety_stock", value)
        End Set
    End Property
    Property Procurement_cycle As Integer
        Get
            Return GoodsQuantity("Procurement_cycle")
        End Get
        Set(value As Integer)
            UpdateGoodsMsg("Procurement_cycle", value)
        End Set
    End Property
    ''' <summary>
    ''' 功能：构造物品
    ''' </summary>
    ''' <param name="SN">序列号</param>
    Private Sub getGoods(ByRef SN As String)
        Sqlcmd.CommandText = "SELECT SN, Department_Type, Nm_Type, Nm FROM ConsumablesManagement_Kc Where SN='" + SN + "'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    Department_Type = dr("Department_Type").ToString
                    GoodsType = dr("Nm_Type").ToString
                    GoodsName = dr("Nm").ToString
                    GoodsSN = SN
                    dr.Close()
                    SqlConn.Close()
                End If
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
        End Try
    End Sub
    ''' <summary>
    ''' 获取物品的SN
    ''' </summary>
    ''' <param name="Department_Type">部门类别</param>
    ''' <param name="GoodsType">物品类别</param>
    ''' <param name="GoodsName">物品名称</param>
    ''' <returns></returns>
    Function getGoodsSN(ByRef Department_Type As String, ByRef GoodsType As String, ByRef GoodsName As String) As String
        Sqlcmd.CommandText = "SELECT SN, Department_Type, Nm_Type, Nm FROM ConsumablesManagement_Kc Where Department_Type='" + Department_Type + "' and Nm_Type='" + GoodsType + "' and Nm='" + GoodsName + "'"
        SqlConn.Open()
        getGoodsSN = ""
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    getGoodsSN = dr("SN").ToString
                    GoodsSN = getGoodsSN
                    dr.Close()
                    SqlConn.Close()
                End If
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
            getGoodsSN = ex.Message
        End Try
        Return getGoodsSN
    End Function
    ''' <summary>
    ''' 功能说明：创建新的耗材
    ''' </summary>
    ''' <param name="Department_Type"></param>
    ''' <param name="GoodsType"></param>
    ''' <param name="GoodsName"></param>
    ''' <returns></returns>
    Function CreatGoods(ByRef Department_Type As String, ByRef GoodsType As String, ByRef GoodsName As String, ByRef UseOperator As String) As String
        If getGoodsSN(Department_Type, GoodsType, GoodsName) = "" Then
            Dim SqlStr As String = ""
            SqlStr = SqlStr & "Insert into ConsumablesManagement_Kc (Department_Type, Nm_Type, Nm,SN,State) Values ('" & Department_Type & "','" & GoodsType & "','" & GoodsName & "','add','启用')"
            SqlStr = SqlStr & Chr(13)
            SqlStr = SqlStr & "Update ConsumablesManagement_Kc Set SN=Department_Type+'A'+right('000000'+cast(id as varchar),5) WHERE Department_Type='" & Department_Type & "' AND Nm_Type='" & GoodsType & "' and Nm='" & GoodsName & "'"
            SqlStr = SqlStr & Chr(13)
            SqlStr = SqlStr & "INSERT INTO ConsumablesManagement_Flow_io "
            SqlStr = SqlStr & " (SN,Department_Type, Nm_Type, Nm,State_Method, Transaction_time, Transaction_person) "
            SqlStr = SqlStr & " Select SN,Department_Type, Nm_Type, Nm,'新增物品','" + Now + "','" & UseOperator & "'"
            SqlStr = SqlStr & " From ConsumablesManagement_Kc WHERE Department_Type='" & Department_Type & "' AND Nm_Type='" & GoodsType & "' and Nm='" & GoodsName & "'"
            Return ToSql(SqlStr, 3)
        Else
            Return "已存在"
        End If
    End Function
    Function DeleteGoods() As String
        If getGoodsSN(Department_Type, GoodsType, GoodsName) = "" Then
            Return "不存在"
        Else
            Dim SqlStr As String = ""
            SqlStr = SqlStr & "INSERT INTO ConsumablesManagement_Flow_io "
            SqlStr = SqlStr & " (SN,Department_Type, Nm_Type, Nm,State_Method, Transaction_time, Transaction_person, State, Purchase_quantity, Safety_stock, Warehouse_stock_quantity, Inventory_quantity_1F, Procurement_cycle, Remarks) "
            SqlStr = SqlStr & " Select SN,Department_Type, Nm_Type, Nm,'删除物品','" + Now + "','" & UseOperator & "', State, Purchase_quantity, Safety_stock, Warehouse_stock_quantity, Inventory_quantity_1F, Procurement_cycle, Remarks"
            SqlStr = SqlStr & " From ConsumablesManagement_Kc  Where SN='" + GoodsSN + "'"
            SqlStr = SqlStr & Chr(13)
            SqlStr = SqlStr & "Delete From ConsumablesManagement_Kc Where SN='" + GoodsSN + "'"
            SqlStr = SqlStr & Chr(13)
            Return ToSql(SqlStr, 2)
        End If
    End Function
    ''' <summary>
    ''' 返回库存表中物品数量
    ''' </summary>
    ''' <param name="SeachType">查询列</param>
    ''' <returns>返回库存表中物品数量</returns>
    Function GoodsQuantity(ByRef SeachType As String) As Integer
        Sqlcmd.CommandText = "SELECT " + SeachType + " FROM ConsumablesManagement_Kc Where SN='" + GoodsSN + "'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    If Not dr(SeachType) Is DBNull.Value Then
                        GoodsQuantity = dr(SeachType)
                    Else
                        GoodsQuantity = 0
                    End If
                    dr.Close()
                    SqlConn.Close()
                End If
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
            GoodsQuantity = 0
        End Try
        Return GoodsQuantity
    End Function
    Function SearchGoodsMsg(ByRef SeachType As String) As String
        Sqlcmd.CommandText = "SELECT " + SeachType + " FROM ConsumablesManagement_Kc Where SN='" + GoodsSN + "'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    If Not dr(SeachType) Is DBNull.Value Then
                        SearchGoodsMsg = dr(SeachType).ToString
                    Else
                        SearchGoodsMsg = ""
                    End If
                    dr.Close()
                    SqlConn.Close()
                Else
                    SearchGoodsMsg = ""
                End If
            Else
                SearchGoodsMsg = ""
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
            SearchGoodsMsg = ""
        End Try
        Return SearchGoodsMsg
    End Function
    Function UpdateGoodsMsg(ByRef Item As String, ByRef Value As String) As String
        Dim SqlStr As String = ""
        ''插入流水记录-新增属性变更记录
        SqlStr = SqlStr & "insert into ConsumablesManagement_Flow_io (SN, Department_Type, Nm_Type, Nm, State_Method, State,Change_Before, Change_After, Transaction_time, Transaction_person)"
        SqlStr = SqlStr & " Select SN, Department_Type, Nm_Type, Nm,'属性变更','" & Item & "'," & Item & ",'" & Value & "','" & Now & "','" & UseOperator & "'"
        SqlStr = SqlStr & "  From ConsumablesManagement_Kc Where SN='" + GoodsSN + "'"
        ''更新库存表-物品属性
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & "update ConsumablesManagement_Kc set " & Item & "= '" & Value & "' Where SN='" + GoodsSN + "'"
        Return ToSql(SqlStr, 2)
    End Function
    ''' <summary>
    ''' 检查返回物品请购数量与入库数量差值（未固定单号）不正确
    ''' </summary>
    ''' <returns></returns>
    Function CheckRequisitionQuantity() As Integer
        Sqlcmd.CommandText = "Select sum(Purchase_quantity) As Num1,sum(Number_of_storage) As Num2 FROM ConsumablesManagement_Flow_io Where SN='" + GoodsSN + "'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    If Not dr("Num1") Is DBNull.Value AndAlso Not dr("Num2") Is DBNull.Value Then
                        CheckRequisitionQuantity = dr("Num1") - dr("Num2")
                    End If
                    dr.Close()
                    SqlConn.Close()
                End If
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
            CheckRequisitionQuantity = 0
        End Try
        Return CheckRequisitionQuantity
    End Function
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
    ''' 新开请购单
    ''' </summary>
    ''' <param name="RequisitionNumber">请购单号</param>
    ''' <param name="PurchaseQuantity">请购数量</param>
    ''' <param name="PurchaseDate">请购日期</param>
    ''' <param name="PurchaseMan">请购人</param>
    ''' <param name="Remarks">备注信息</param>
    ''' <returns>返回执行结果</returns>
    Function PurchaseRequisition_New(ByRef RequisitionNumber As String, ByRef PurchaseQuantity As Integer, ByRef PurchaseDate As Date, ByRef PurchaseMan As String, ByRef Remarks As String) As String
        Dim TotalPurchaseQuantity As Integer = GoodsQuantity("Purchase_quantity") + PurchaseQuantity
        SqlStr = ""
        ''更新流水记录-新增请购单
        SqlStr = SqlStr & "INSERT INTO ConsumablesManagement_Flow_io "
        SqlStr = SqlStr & "(SN,Department_Type, Nm_Type, Nm, Requisition_number, Purchase_quantity, Number_of_storage, "
        SqlStr = SqlStr & "Remarks,Date_of_purchase, State_Method, State, Transaction_time, Transaction_person ) "
        SqlStr = SqlStr & " Select SN,Department_Type, Nm_Type, Nm,'" & RequisitionNumber & "','" & PurchaseQuantity & "',0,"
        SqlStr = SqlStr & " '" + Remarks + "','" + PurchaseDate + "','新增请购','未结','" + Now + "','" + PurchaseMan + "'"
        SqlStr = SqlStr & " From ConsumablesManagement_Kc Where SN='" + GoodsSN + "'"
        SqlStr = SqlStr & Chr(13)
        ''更新流水记录-更新标识ID
        SqlStr = SqlStr & " UPDATE ConsumablesManagement_Flow_io set Id_number=(Requisition_number  + '_' +  convert(nvarchar(10),Id) )  Where State_Method='新增请购' and  State='未结' and Id_number is NULL and SN='" + GoodsSN + "'"
        SqlStr = SqlStr & Chr(13)
        ''更新库存请购数量
        SqlStr = SqlStr & " UPDATE ConsumablesManagement_Kc set Purchase_quantity='" & TotalPurchaseQuantity & "'  Where SN='" + GoodsSN + "'"
        Return ToSql(SqlStr, 3)
    End Function
    ''' <summary>
    ''' 入库
    ''' </summary>
    ''' <param name="Location">入库位置</param>
    ''' <param name="AcceptanceQuantity">入库数量</param>
    ''' <param name="AcceptanceMan">入库人员</param>
    ''' <param name="IdNumber">请购单号+序号</param>
    ''' <returns>返回执行结果</returns>
    Function PurchaseAcceptance(ByRef Location As String, ByRef AcceptanceQuantity As Integer, ByRef AcceptanceMan As String, ByRef IdNumber As String) As String
        Dim NewQuantity As Integer = GoodsQuantity(Location) + AcceptanceQuantity
        If CheckAcceptancQuantity(AcceptanceQuantity, IdNumber) Then
            SqlStr = ""
            ''插入流水记录-新增入库记录
            SqlStr = SqlStr & "INSERT INTO ConsumablesManagement_Flow_io  "
            SqlStr = SqlStr & "( Id_number,SN, Department_Type, Nm_Type, Nm, Requisition_number, Purchase_quantity, Number_of_storage, "
            SqlStr = SqlStr & " Storage_location, "
            SqlStr = SqlStr & " State_Method, Transaction_time, Transaction_person," & Location & " )"
            SqlStr = SqlStr & " Select Id_number,SN, Department_Type, Nm_Type, Nm, Requisition_number, Purchase_quantity,'" & AcceptanceQuantity & "',"
            SqlStr = SqlStr & "  (case when '" + Location + "'='Warehouse_stock_quantity' then '仓库' else '1F无尘室' end),"
            SqlStr = SqlStr & " '新增入库','" + Now + "','" + AcceptanceMan + "','" & AcceptanceQuantity & "'"
            SqlStr = SqlStr & " From ConsumablesManagement_Flow_io Where  Id_number='" & IdNumber & "' and State_Method='新增请购'"
            ''更新流水记录-更新原请购单入库数量
            SqlStr = SqlStr & Chr(13)
            SqlStr = SqlStr & "Update ConsumablesManagement_Flow_io set "
            SqlStr = SqlStr & " " & Location & "=ISNULL(" & Location & ",0)+'" & AcceptanceQuantity & "'"
            SqlStr = SqlStr & "  ,Number_of_storage=Number_of_storage+'" & AcceptanceQuantity & "'"
            SqlStr = SqlStr & "  , State=(CASE WHEN Purchase_quantity=Number_of_storage+'" & AcceptanceQuantity & "' THEN '已结' Else '未结' end)"
            SqlStr = SqlStr & "  Where  Id_number='" & IdNumber & "' and State_Method='新增请购'"
            ''更新库存表-物品请购数量，入库位置数量
            SqlStr = SqlStr & Chr(13)
            SqlStr = SqlStr & "Update ConsumablesManagement_Kc set "
            SqlStr = SqlStr & " Purchase_quantity=Purchase_quantity - '" & AcceptanceQuantity & "'," & Location & "='" & NewQuantity & "'"
            SqlStr = SqlStr & "  Where SN='" + GoodsSN + "'"
            Return ToSql(SqlStr, 3)
        Else
            Return "新入库数量超出当前剩余请购单数量超出"
        End If
    End Function
    ''' <summary>
    ''' 领用
    ''' </summary>
    ''' <param name="Quantity">领用数量</param>
    ''' <param name="Msg">信息集合</param>
    ''' <param name="UseOperator">操作人员</param>
    ''' <returns></returns>
    Function Registration(ByRef Quantity As Integer, ByRef Msg As String, ByRef UseOperator As String) As String
        Dim Use_Dt As String = Split(Msg, "%")(0)
        Dim Use_area As String = Split(Msg, "%")(1)
        Dim Purpose As String = Split(Msg, "%")(2)
        Dim Receive_No As String = UCase(Split(Msg, "%")(3).Trim)
        Dim Remarks As String = Split(Msg, "%")(4)
        Dim p As New JwrEmployee(Receive_No)
        Dim Receive_Person As String = p.Name
        Dim Receive_Person_Station As String = p.Station
        p = Nothing
        SqlStr = ""
        ''插入流水记录-新增领用记录
        SqlStr = SqlStr & "INSERT INTO ConsumablesManagement_Flow_io  ( Department_Type,SN, Nm_Type,Nm, State_Method, Receive_quantity, Receive_time "
        SqlStr = SqlStr & " ,Receive_No,Receive_Person, Receive_Person_Station, Transaction_time"
        SqlStr = SqlStr & " , Transaction_person,Use_Dt, Use_area, Purpose,Remarks,Inventory_quantity_1F, Inventory_quantity_1F_new )"
        SqlStr = SqlStr & " Select Department_Type,SN, Nm_Type,Nm,'新增领用','" & Quantity & "','" + Now + "'"
        SqlStr = SqlStr & " ,'" & Receive_No & "','" & Receive_Person & "','" & Receive_Person_Station & "','" + Now + "'"
        SqlStr = SqlStr & " ,'" & UseOperator & "','" & Use_Dt & "','" & Use_area & "','" & Purpose & "','" & Remarks & "',Inventory_quantity_1F,Inventory_quantity_1F-'" & Quantity & "'"
        SqlStr = SqlStr & " From ConsumablesManagement_Kc Where SN='" + GoodsSN + "'"
        ''更新库存表-1F无尘室物品数量
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & "UPDATE ConsumablesManagement_Kc Set Inventory_quantity_1F=Inventory_quantity_1F-'" & Quantity & "' where SN='" + GoodsSN + "'"
        Return ToSql(SqlStr, 2)
    End Function

    Function Dump(ByRef Remarks As String, ByRef NewQuantity_Dump As Integer, ByRef NewQuantity_1F As Integer, ByRef NewQuantity_Warehouse As Integer, ByRef UseOperator As String) As String
        SqlStr = ""
        ''插入流水记录-新增转存记录
        SqlStr = SqlStr & "INSERT INTO  ConsumablesManagement_Flow_io  (  Department_Type,SN, Nm_Type, Nm, State_Method, Warehouse_stock_quantity"
        SqlStr = SqlStr & " , Inventory_quantity_1F, Warehouse_stock_quantity_new, Inventory_quantity_1F_new,Dump1,Remarks"
        SqlStr = SqlStr & " , Transaction_time, Transaction_person )"
        SqlStr = SqlStr & "  Select Department_Type,SN, Nm_Type, Nm,'新增转仓', Warehouse_stock_quantity"
        SqlStr = SqlStr & " , Inventory_quantity_1F,'" & NewQuantity_Warehouse & "','" & NewQuantity_1F & "','" & NewQuantity_Dump & "','" & Remarks & "'"
        SqlStr = SqlStr & " ,'" + Now + "','" & UseOperator & "'"
        SqlStr = SqlStr & " From ConsumablesManagement_Kc Where SN='" + GoodsSN + "'"
        ''更新库存表-无尘室和仓库数量
        SqlStr = SqlStr & Chr(13)
        SqlStr = SqlStr & "UPDATE ConsumablesManagement_Kc Set Warehouse_stock_quantity='" & NewQuantity_Warehouse & "',Inventory_quantity_1F='" & NewQuantity_1F & "' Where SN='" + GoodsSN + "'"
        Return ToSql(SqlStr, 2)
    End Function
    ''' <summary>
    ''' 检查新入库数量是否超过当前剩余请购单数量
    ''' </summary>
    ''' <param name="AcceptanceQuantity">入库数量</param>
    ''' <param name="IdNumber">请购单号+序号</param>
    ''' <returns></returns>
    Function CheckAcceptancQuantity(ByRef AcceptanceQuantity As Integer, ByRef IdNumber As String) As Boolean
        Sqlcmd.CommandText = "Select Purchase_quantity-Number_of_storage As Num FROM ConsumablesManagement_Flow_io Where Id_number ='" & IdNumber & "' and SN='" + GoodsSN + "' and State_Method='新增请购'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    If Val(dr("Num")) - AcceptanceQuantity >= 0 Then
                        CheckAcceptancQuantity = True
                    End If
                    dr.Close()
                    SqlConn.Close()
                End If
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
        End Try
        Return CheckAcceptancQuantity
    End Function

    ''' <summary>
    ''' 获取Flow表中该序号对应的请购单号
    ''' </summary>
    ''' <param name="Id">Flow序号</param>
    ''' <returns>请购单号+序号</returns>
    Private Function getFlowIdNumber(ByRef Id As Integer) As String
        Sqlcmd.CommandText = "SELECT Id_number FROM ConsumablesManagement_Flow_io Where Id='" & Id & "' and SN='" + GoodsSN + "'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    If Not dr("Id_number") Is DBNull.Value Then
                        getFlowIdNumber = dr("Id_number").ToString
                    Else
                        Return ""
                    End If
                    dr.Close()
                    SqlConn.Close()
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
            Return ""
        End Try
        Return getFlowIdNumber
    End Function
    ''' <summary>
    ''' 获取入库记录中，该请购单总入库量
    ''' </summary>
    ''' <param name="Id">Flow序号</param>
    ''' <returns></returns>
    Function getAcceptancQuantity(ByRef Id As Integer) As Integer
        Sqlcmd.CommandText = "SELECT Sum(Number_of_storage) as Num1 FROM ConsumablesManagement_Flow_io Where  SN='" + GoodsSN + "' and  State_Method='新增入库'"
        SqlConn.Open()
        Try
            dr = Sqlcmd.ExecuteReader
            If dr.HasRows Then
                If dr.Read Then
                    If Not dr("Num1") Is DBNull.Value Then
                        getAcceptancQuantity = dr("Num1")
                    End If
                    dr.Close()
                    SqlConn.Close()
                End If
            End If
            dr.Close()
            SqlConn.Close()
        Catch ex As Exception
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            SqlConn.Close()
        End Try
        Return getAcceptancQuantity
    End Function
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
