Imports System.IO
Imports Microsoft.VisualBasic

Public Class ProgressBar
    Inherits System.Web.UI.Page
    Private m_page As System.Web.UI.Page
    ''' <summary>
    ''' 最大值
    ''' </summary>
    ''' <returns></returns>
    Property MaxValue As Integer
        Get
            If IsDBNull(ViewState("MaxValue")) Then
                Return 0
            Else
                Return Convert.ToInt32(ViewState("MaxValue"))
            End If
        End Get
        Set(value As Integer)
            ViewState("MaxValue") = value
        End Set
    End Property
    ''' <summary>
    ''' 当前值
    ''' </summary>
    ''' <returns></returns>
    Property ThisValue As Integer
        Get
            If IsDBNull(ViewState("ThisValue")) Then
                Return 0
            Else
                Return Convert.ToInt32(ViewState("ThisValue"))
            End If
        End Get
        Set(value As Integer)
            ViewState("ThisValue") = value
        End Set
    End Property
    ''' <summary>
    ''' 功能描述：构造函数
    ''' </summary>
    ''' <param name="page">当前页面</param>
    Public Sub New(ByRef page As System.Web.UI.Page)
        m_page = page
    End Sub
    Public Sub SetMaxValue(ByRef intMaxValue As Integer)
        MaxValue = intMaxValue
    End Sub
    ''' <summary>
    ''' 功能描述:初始化进度条
    ''' </summary>
    Public Sub InitProgress()
        Dim templateFileName As String = AppDomain.CurrentDomain.BaseDirectory + "ProgressBar.htm"
        Dim reader As StreamReader = New StreamReader(templateFileName, System.Text.Encoding.GetEncoding("GB2312"))
        Dim strhtml As String = reader.ReadToEnd()
        reader.Close()
        m_page.Response.Write(strhtml)
        m_page.Response.Flush()
    End Sub
    ''' <summary>
    ''' 功能描述:设置标题
    ''' </summary>
    ''' <param name="strTitle">strTitle</param>
    Public Sub SetTitle(ByRef strTitle As String)
        Dim strjsBlock As String = "<script>SetTitle('" + strTitle + "'); </script>"
        m_page.Response.Write(strjsBlock)
        m_page.Response.Flush()
    End Sub
    ''' <summary>
    ''' 功能描述:设置进度
    ''' </summary>
    ''' <param name="intpercent"></param>
    Public Sub AddProgress(ByRef intpercent As Integer)
        ThisValue = ThisValue + intpercent
        Dim dblstep As Double = (CType(ThisValue, Double) / CType(MaxValue, Double)) * 100
        Dim strjsBlock As String = "<script>SetPorgressBar('" + dblstep.ToString("0.00") + "'); </script>"
        m_page.Response.Write(strjsBlock)
        m_page.Response.Flush()
    End Sub
    Public Sub DisponseProgress()
        Dim strjsBlock As String = "<script>SetCompleted();</script>"
        m_page.Response.Write(strjsBlock)
        m_page.Response.Flush()
    End Sub
End Class
