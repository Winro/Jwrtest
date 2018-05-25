Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class Material

    Dim SqlConn As New SqlConnection(ConfigurationManager.ConnectionStrings("NB410Sys").ConnectionString)
    Dim Sqlcmd As SqlCommand = SqlConn.CreateCommand
    Dim Trans As SqlTransaction
    Dim dr As SqlDataReader
    Dim SqlStr As String = ""
    Private Department_Type As String
    Private GoodsType As String
    Private GoodsName As String
    Private GoodsSN As String
    Private Nm As String
    Private NmType As String
    Private UseDepartment As String
    Private State As String
    Private Remarks As String
    Private Space_Number As String
    Private Space_Type As String
    Dim S_x As Integer
    Dim S_y As Integer
    Dim S_w As Integer
    Dim S_h As Integer
    Private Purchase_quantity As Integer
    Private Safety_stock As Integer
    Private Warehouse_stock_quantity As Integer
    Private Inventory_quantity_1F As Integer
    Private Procurement_cycle As Integer
    Private Length As Single
    Private Wideth As Single

    Property SetMaterialName() As String
        Get
            Return Nm
        End Get
        Set(value As String)
            Nm = value
        End Set
    End Property
    Function CreatMaterial() As Integer
        Dim i As Integer
        SqlConn.Open()
        Sqlcmd.CommandText = "insert into TBL (tt) values ('" + Nm + "')"
        i = Sqlcmd.ExecuteNonQuery()
        SqlConn.Close()
        Return i
    End Function
    Private Function MCreatItem() As String
        Dim SqlItem As String = ""
        Select Case NmType
            Case ""
            Case ""
        End Select
        Return SqlItem
    End Function
End Class
