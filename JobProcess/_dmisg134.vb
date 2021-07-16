Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.DMISG
  <DataObject()> _
  Partial Public Class dmisg134
    Private Shared _RecordCount As Integer
    Private _t_rcno As String = ""
    Private _t_revn As String = ""
    Private _t_cprj As String = ""
    Private _t_item As String = ""
    Private _t_bpid As String = ""
    Private _t_nama As String = ""
    Private _t_stat As Int32 = 0
    Private _t_user As String = ""
    Private _t_date As String = ""
    Private _t_sent_1 As Int32 = 0
    Private _t_sent_2 As Int32 = 0
    Private _t_sent_3 As Int32 = 0
    Private _t_sent_4 As Int32 = 0
    Private _t_sent_5 As Int32 = 0
    Private _t_sent_6 As Int32 = 0
    Private _t_sent_7 As Int32 = 0
    Private _t_rece_1 As Int32 = 0
    Private _t_rece_2 As Int32 = 0
    Private _t_rece_3 As Int32 = 0
    Private _t_rece_4 As Int32 = 0
    Private _t_rece_5 As Int32 = 0
    Private _t_rece_6 As Int32 = 0
    Private _t_rece_7 As Int32 = 0
    Private _t_suer As String = ""
    Private _t_sdat As String = ""
    Private _t_appr As String = ""
    Private _t_adat As String = ""
    Private _t_subm_1 As Int32 = 0
    Private _t_subm_2 As Int32 = 0
    Private _t_subm_3 As Int32 = 0
    Private _t_subm_4 As Int32 = 0
    Private _t_subm_5 As Int32 = 0
    Private _t_subm_6 As Int32 = 0
    Private _t_subm_7 As Int32 = 0
    Private _t_orno As String = ""
    Private _t_pono As Int32 = 0
    Private _t_trno As String = ""
    Private _t_Refcntd As Int32 = 0
    Private _t_Refcntu As Int32 = 0
    Private _t_docn As String = ""
    Private _t_eunt As String = ""
    Private _t_atch As String = ""
    Private _t_rqln As Int32 = 0
    Private _t_rqno As String = ""
    Private _t_pwfd As Int32 = 0
    Private _t_wfid As Int32 = 0
    Private _t_apid_1 As String = ""
    Private _t_apid_2 As String = ""
    Private _t_apid_3 As String = ""
    Private _t_apid_4 As String = ""
    Private _t_apid_5 As String = ""
    Private _t_apid_6 As String = ""
    Private _t_apid_7 As String = ""
    Public Property t_rcno() As String
      Get
        Return _t_rcno
      End Get
      Set(ByVal value As String)
        _t_rcno = value
      End Set
    End Property
    Public Property t_revn() As String
      Get
        Return _t_revn
      End Get
      Set(ByVal value As String)
        _t_revn = value
      End Set
    End Property
    Public Property t_cprj() As String
      Get
        Return _t_cprj
      End Get
      Set(ByVal value As String)
        _t_cprj = value
      End Set
    End Property
    Public Property t_item() As String
      Get
        Return _t_item
      End Get
      Set(ByVal value As String)
        _t_item = value
      End Set
    End Property
    Public Property t_bpid() As String
      Get
        Return _t_bpid
      End Get
      Set(ByVal value As String)
        _t_bpid = value
      End Set
    End Property
    Public Property t_nama() As String
      Get
        Return _t_nama
      End Get
      Set(ByVal value As String)
        _t_nama = value
      End Set
    End Property
    Public Property t_stat() As Int32
      Get
        Return _t_stat
      End Get
      Set(ByVal value As Int32)
        _t_stat = value
      End Set
    End Property
    Public Property t_user() As String
      Get
        Return _t_user
      End Get
      Set(ByVal value As String)
        _t_user = value
      End Set
    End Property
    Public Property t_date() As String
      Get
        If Not _t_date = String.Empty Then
          Return Convert.ToDateTime(_t_date).ToString("dd/MM/yyyy")
        End If
        Return _t_date
      End Get
      Set(ByVal value As String)
         _t_date = value
      End Set
    End Property
    Public Property t_sent_1() As Int32
      Get
        Return _t_sent_1
      End Get
      Set(ByVal value As Int32)
        _t_sent_1 = value
      End Set
    End Property
    Public Property t_sent_2() As Int32
      Get
        Return _t_sent_2
      End Get
      Set(ByVal value As Int32)
        _t_sent_2 = value
      End Set
    End Property
    Public Property t_sent_3() As Int32
      Get
        Return _t_sent_3
      End Get
      Set(ByVal value As Int32)
        _t_sent_3 = value
      End Set
    End Property
    Public Property t_sent_4() As Int32
      Get
        Return _t_sent_4
      End Get
      Set(ByVal value As Int32)
        _t_sent_4 = value
      End Set
    End Property
    Public Property t_sent_5() As Int32
      Get
        Return _t_sent_5
      End Get
      Set(ByVal value As Int32)
        _t_sent_5 = value
      End Set
    End Property
    Public Property t_sent_6() As Int32
      Get
        Return _t_sent_6
      End Get
      Set(ByVal value As Int32)
        _t_sent_6 = value
      End Set
    End Property
    Public Property t_sent_7() As Int32
      Get
        Return _t_sent_7
      End Get
      Set(ByVal value As Int32)
        _t_sent_7 = value
      End Set
    End Property
    Public Property t_rece_1() As Int32
      Get
        Return _t_rece_1
      End Get
      Set(ByVal value As Int32)
        _t_rece_1 = value
      End Set
    End Property
    Public Property t_rece_2() As Int32
      Get
        Return _t_rece_2
      End Get
      Set(ByVal value As Int32)
        _t_rece_2 = value
      End Set
    End Property
    Public Property t_rece_3() As Int32
      Get
        Return _t_rece_3
      End Get
      Set(ByVal value As Int32)
        _t_rece_3 = value
      End Set
    End Property
    Public Property t_rece_4() As Int32
      Get
        Return _t_rece_4
      End Get
      Set(ByVal value As Int32)
        _t_rece_4 = value
      End Set
    End Property
    Public Property t_rece_5() As Int32
      Get
        Return _t_rece_5
      End Get
      Set(ByVal value As Int32)
        _t_rece_5 = value
      End Set
    End Property
    Public Property t_rece_6() As Int32
      Get
        Return _t_rece_6
      End Get
      Set(ByVal value As Int32)
        _t_rece_6 = value
      End Set
    End Property
    Public Property t_rece_7() As Int32
      Get
        Return _t_rece_7
      End Get
      Set(ByVal value As Int32)
        _t_rece_7 = value
      End Set
    End Property
    Public Property t_suer() As String
      Get
        Return _t_suer
      End Get
      Set(ByVal value As String)
        _t_suer = value
      End Set
    End Property
    Public Property t_sdat() As String
      Get
        If Not _t_sdat = String.Empty Then
          Return Convert.ToDateTime(_t_sdat).ToString("dd/MM/yyyy")
        End If
        Return _t_sdat
      End Get
      Set(ByVal value As String)
         _t_sdat = value
      End Set
    End Property
    Public Property t_appr() As String
      Get
        Return _t_appr
      End Get
      Set(ByVal value As String)
        _t_appr = value
      End Set
    End Property
    Public Property t_adat() As String
      Get
        If Not _t_adat = String.Empty Then
          Return Convert.ToDateTime(_t_adat).ToString("dd/MM/yyyy")
        End If
        Return _t_adat
      End Get
      Set(ByVal value As String)
         _t_adat = value
      End Set
    End Property
    Public Property t_subm_1() As Int32
      Get
        Return _t_subm_1
      End Get
      Set(ByVal value As Int32)
        _t_subm_1 = value
      End Set
    End Property
    Public Property t_subm_2() As Int32
      Get
        Return _t_subm_2
      End Get
      Set(ByVal value As Int32)
        _t_subm_2 = value
      End Set
    End Property
    Public Property t_subm_3() As Int32
      Get
        Return _t_subm_3
      End Get
      Set(ByVal value As Int32)
        _t_subm_3 = value
      End Set
    End Property
    Public Property t_subm_4() As Int32
      Get
        Return _t_subm_4
      End Get
      Set(ByVal value As Int32)
        _t_subm_4 = value
      End Set
    End Property
    Public Property t_subm_5() As Int32
      Get
        Return _t_subm_5
      End Get
      Set(ByVal value As Int32)
        _t_subm_5 = value
      End Set
    End Property
    Public Property t_subm_6() As Int32
      Get
        Return _t_subm_6
      End Get
      Set(ByVal value As Int32)
        _t_subm_6 = value
      End Set
    End Property
    Public Property t_subm_7() As Int32
      Get
        Return _t_subm_7
      End Get
      Set(ByVal value As Int32)
        _t_subm_7 = value
      End Set
    End Property
    Public Property t_orno() As String
      Get
        Return _t_orno
      End Get
      Set(ByVal value As String)
        _t_orno = value
      End Set
    End Property
    Public Property t_pono() As Int32
      Get
        Return _t_pono
      End Get
      Set(ByVal value As Int32)
        _t_pono = value
      End Set
    End Property
    Public Property t_trno() As String
      Get
        Return _t_trno
      End Get
      Set(ByVal value As String)
        _t_trno = value
      End Set
    End Property
    Public Property t_Refcntd() As Int32
      Get
        Return _t_Refcntd
      End Get
      Set(ByVal value As Int32)
        _t_Refcntd = value
      End Set
    End Property
    Public Property t_Refcntu() As Int32
      Get
        Return _t_Refcntu
      End Get
      Set(ByVal value As Int32)
        _t_Refcntu = value
      End Set
    End Property
    Public Property t_docn() As String
      Get
        Return _t_docn
      End Get
      Set(ByVal value As String)
        _t_docn = value
      End Set
    End Property
    Public Property t_eunt() As String
      Get
        Return _t_eunt
      End Get
      Set(ByVal value As String)
        _t_eunt = value
      End Set
    End Property
    Public Property t_atch() As String
      Get
        Return _t_atch
      End Get
      Set(ByVal value As String)
        _t_atch = value
      End Set
    End Property
    Public Property t_rqln() As Int32
      Get
        Return _t_rqln
      End Get
      Set(ByVal value As Int32)
        _t_rqln = value
      End Set
    End Property
    Public Property t_rqno() As String
      Get
        Return _t_rqno
      End Get
      Set(ByVal value As String)
        _t_rqno = value
      End Set
    End Property
    Public Property t_pwfd() As Int32
      Get
        Return _t_pwfd
      End Get
      Set(ByVal value As Int32)
        _t_pwfd = value
      End Set
    End Property
    Public Property t_wfid() As Int32
      Get
        Return _t_wfid
      End Get
      Set(ByVal value As Int32)
        _t_wfid = value
      End Set
    End Property
    Public Property t_apid_1() As String
      Get
        Return _t_apid_1
      End Get
      Set(ByVal value As String)
        _t_apid_1 = value
      End Set
    End Property
    Public Property t_apid_2() As String
      Get
        Return _t_apid_2
      End Get
      Set(ByVal value As String)
        _t_apid_2 = value
      End Set
    End Property
    Public Property t_apid_3() As String
      Get
        Return _t_apid_3
      End Get
      Set(ByVal value As String)
        _t_apid_3 = value
      End Set
    End Property
    Public Property t_apid_4() As String
      Get
        Return _t_apid_4
      End Get
      Set(ByVal value As String)
        _t_apid_4 = value
      End Set
    End Property
    Public Property t_apid_5() As String
      Get
        Return _t_apid_5
      End Get
      Set(ByVal value As String)
        _t_apid_5 = value
      End Set
    End Property
    Public Property t_apid_6() As String
      Get
        Return _t_apid_6
      End Get
      Set(ByVal value As String)
        _t_apid_6 = value
      End Set
    End Property
    Public Property t_apid_7() As String
      Get
        Return _t_apid_7
      End Get
      Set(ByVal value As String)
        _t_apid_7 = value
      End Set
    End Property
    Public Readonly Property DisplayField() As String
      Get
        Return ""
      End Get
    End Property
    Public Readonly Property PrimaryKey() As String
      Get
        Return _t_rcno & "|" & _t_revn
      End Get
    End Property
    Public Shared Property RecordCount() As Integer
      Get
        Return _RecordCount
      End Get
      Set(ByVal value As Integer)
        _RecordCount = value
      End Set
    End Property
    Public Class PKdmisg134
      Private _t_rcno As String = ""
      Private _t_revn As String = ""
      Public Property t_rcno() As String
        Get
          Return _t_rcno
        End Get
        Set(ByVal value As String)
          _t_rcno = value
        End Set
      End Property
      Public Property t_revn() As String
        Get
          Return _t_revn
        End Get
        Set(ByVal value As String)
          _t_revn = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function dmisg134GetNewRecord() As SIS.DMISG.dmisg134
      Return New SIS.DMISG.dmisg134()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function dmisg134GetByID(ByVal t_rcno As String, ByVal t_revn As String) As SIS.DMISG.dmisg134
      Dim Results As SIS.DMISG.dmisg134 = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "spdmisg134SelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_rcno", SqlDbType.VarChar, t_rcno.ToString.Length, t_rcno)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@t_revn", SqlDbType.VarChar, t_revn.ToString.Length, t_revn)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.DMISG.dmisg134(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function CommentSubmittedReceipts(ByVal SinceDays As Integer) As List(Of SIS.DMISG.dmisg134)
      Dim Sql As String = ""
      Sql &= " select * from tdmisg134200 "
      Sql &= " where left(t_rcno,3) = 'REC' " 'Preorder Receipts
      Sql &= " and t_wfid > 0 " 'Created From online portal
      Sql &= " and t_stat IN (4,5) " 'Commentsubmitted or Technically Cleared
      Sql &= " and t_adat > convert(datetime,'" & Now.AddDays(-1 * SinceDays).ToString("dd/MM/yyyy") & "',103) " 'Receipts of after date
      Dim Results As List(Of SIS.DMISG.dmisg134) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = Sql
          Results = New List(Of SIS.DMISG.dmisg134)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.DMISG.dmisg134(Reader))
          End While
          Reader.Close()
          _RecordCount = Results.Count
        End Using
      End Using
      Return Results
    End Function

    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function dmisg134SelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String) As List(Of SIS.DMISG.dmisg134)
      Dim Results As List(Of SIS.DMISG.dmisg134) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "spdmisg134SelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "spdmisg134SelectListFilteres"
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.DMISG.dmisg134)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.DMISG.dmisg134(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function dmisg134SelectCount(ByVal SearchState As Boolean, ByVal SearchText As String) As Integer
      Return _RecordCount
    End Function
    'Select By ID One Record Filtered Overloaded GetByID 
    Public Sub New(ByVal Reader As SqlDataReader)
      Try
        For Each pi As System.Reflection.PropertyInfo In Me.GetType.GetProperties
          If pi.MemberType = Reflection.MemberTypes.Property Then
            Try
              Dim Found As Boolean = False
              For I As Integer = 0 To Reader.FieldCount - 1
                If Reader.GetName(I).ToLower = pi.Name.ToLower Then
                  Found = True
                  Exit For
                End If
              Next
              If Found Then
                If Convert.IsDBNull(Reader(pi.Name)) Then
                  Select Case Reader.GetDataTypeName(Reader.GetOrdinal(pi.Name))
                    Case "decimal"
                      CallByName(Me, pi.Name, CallType.Let, "0.00")
                    Case "bit"
                      CallByName(Me, pi.Name, CallType.Let, Boolean.FalseString)
                    Case Else
                      CallByName(Me, pi.Name, CallType.Let, String.Empty)
                  End Select
                Else
                  CallByName(Me, pi.Name, CallType.Let, Reader(pi.Name))
                End If
              End If
            Catch ex As Exception
            End Try
          End If
        Next
      Catch ex As Exception
      End Try
    End Sub
    Public Sub New()
    End Sub
  End Class
End Namespace
