Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.ERP
  <DataObject()> _
  Partial Public Class erpIDMSAlertSent
    Private Shared _RecordCount As Integer
    Private _ReceiptNo As String = ""
    Private _RevisionNo As String = ""
    Private _MailSentOn As String = ""
    Public Property ReceiptNo() As String
      Get
        Return _ReceiptNo
      End Get
      Set(ByVal value As String)
        _ReceiptNo = value
      End Set
    End Property
    Public Property RevisionNo() As String
      Get
        Return _RevisionNo
      End Get
      Set(ByVal value As String)
        _RevisionNo = value
      End Set
    End Property
    Public Property MailSentOn() As String
      Get
        If Not _MailSentOn = String.Empty Then
          Return Convert.ToDateTime(_MailSentOn).ToString("dd/MM/yyyy")
        End If
        Return _MailSentOn
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _MailSentOn = ""
         Else
           _MailSentOn = value
         End If
      End Set
    End Property
    Public Readonly Property DisplayField() As String
      Get
        Return ""
      End Get
    End Property
    Public Readonly Property PrimaryKey() As String
      Get
        Return _ReceiptNo & "|" & _RevisionNo
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
    Public Class PKerpIDMSAlertSent
      Private _ReceiptNo As String = ""
      Private _RevisionNo As String = ""
      Public Property ReceiptNo() As String
        Get
          Return _ReceiptNo
        End Get
        Set(ByVal value As String)
          _ReceiptNo = value
        End Set
      End Property
      Public Property RevisionNo() As String
        Get
          Return _RevisionNo
        End Get
        Set(ByVal value As String)
          _RevisionNo = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function erpIDMSAlertSentGetNewRecord() As SIS.ERP.erpIDMSAlertSent
      Return New SIS.ERP.erpIDMSAlertSent()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function erpIDMSAlertSentGetByID(ByVal ReceiptNo As String, ByVal RevisionNo As String) As SIS.ERP.erpIDMSAlertSent
      Dim Results As SIS.ERP.erpIDMSAlertSent = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sperpIDMSAlertSentSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ReceiptNo", SqlDbType.NVarChar, ReceiptNo.ToString.Length, ReceiptNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@RevisionNo", SqlDbType.NVarChar, RevisionNo.ToString.Length, RevisionNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.ERP.erpIDMSAlertSent(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)>
    Public Shared Function erpIDMSAlertSentSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String) As List(Of SIS.ERP.erpIDMSAlertSent)
      Dim Results As List(Of SIS.ERP.erpIDMSAlertSent) = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          If SearchState Then
            Cmd.CommandText = "sperpIDMSAlertSentSelectListSearch"
            EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
          Else
            Cmd.CommandText = "sperpIDMSAlertSentSelectListFilteres"
          End If
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
          Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
          Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Results = New List(Of SIS.ERP.erpIDMSAlertSent)()
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          While (Reader.Read())
            Results.Add(New SIS.ERP.erpIDMSAlertSent(Reader))
          End While
          Reader.Close()
          _RecordCount = Cmd.Parameters("@RecordCount").Value
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function erpIDMSAlertSentSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String) As Integer
      Return _RecordCount
    End Function
    'Select By ID One Record Filtered Overloaded GetByID
    <DataObjectMethod(DataObjectMethodType.Insert, True)>
    Public Shared Function erpIDMSAlertSentInsert(ByVal Record As SIS.ERP.erpIDMSAlertSent) As SIS.ERP.erpIDMSAlertSent
      Dim _Rec As SIS.ERP.erpIDMSAlertSent = SIS.ERP.erpIDMSAlertSent.erpIDMSAlertSentGetNewRecord()
      With _Rec
        .ReceiptNo = Record.ReceiptNo
        .RevisionNo = Record.RevisionNo
        .MailSentOn = Record.MailSentOn
      End With
      Return SIS.ERP.erpIDMSAlertSent.InsertData(_Rec)
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.ERP.erpIDMSAlertSent) As SIS.ERP.erpIDMSAlertSent
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sperpIDMSAlertSentInsert"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ReceiptNo", SqlDbType.NVarChar, 10, Record.ReceiptNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@RevisionNo", SqlDbType.NVarChar, 3, Record.RevisionNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MailSentOn", SqlDbType.DateTime, 21, IIf(Record.MailSentOn = "", Convert.DBNull, Record.MailSentOn))
          Cmd.Parameters.Add("@Return_ReceiptNo", SqlDbType.NVarChar, 10)
          Cmd.Parameters("@Return_ReceiptNo").Direction = ParameterDirection.Output
          Cmd.Parameters.Add("@Return_RevisionNo", SqlDbType.NVarChar, 3)
          Cmd.Parameters("@Return_RevisionNo").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.ReceiptNo = Cmd.Parameters("@Return_ReceiptNo").Value
          Record.RevisionNo = Cmd.Parameters("@Return_RevisionNo").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Update, True)>
    Public Shared Function erpIDMSAlertSentUpdate(ByVal Record As SIS.ERP.erpIDMSAlertSent) As SIS.ERP.erpIDMSAlertSent
      Dim _Rec As SIS.ERP.erpIDMSAlertSent = SIS.ERP.erpIDMSAlertSent.erpIDMSAlertSentGetByID(Record.ReceiptNo, Record.RevisionNo)
      With _Rec
        .MailSentOn = Record.MailSentOn
      End With
      Return SIS.ERP.erpIDMSAlertSent.UpdateData(_Rec)
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.ERP.erpIDMSAlertSent) As SIS.ERP.erpIDMSAlertSent
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sperpIDMSAlertSentUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_ReceiptNo", SqlDbType.NVarChar, 10, Record.ReceiptNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_RevisionNo", SqlDbType.NVarChar, 3, Record.RevisionNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ReceiptNo", SqlDbType.NVarChar, 10, Record.ReceiptNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@RevisionNo", SqlDbType.NVarChar, 3, Record.RevisionNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@MailSentOn", SqlDbType.DateTime, 21, IIf(Record.MailSentOn = "", Convert.DBNull, Record.MailSentOn))
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Con.Open()
          Cmd.ExecuteNonQuery()
          _RecordCount = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return Record
    End Function
    <DataObjectMethod(DataObjectMethodType.Delete, True)>
    Public Shared Function erpIDMSAlertSentDelete(ByVal Record As SIS.ERP.erpIDMSAlertSent) As Int32
      Dim _Result As Integer = 0
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sperpIDMSAlertSentDelete"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_ReceiptNo", SqlDbType.NVarChar, Record.ReceiptNo.ToString.Length, Record.ReceiptNo)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_RevisionNo", SqlDbType.NVarChar, Record.RevisionNo.ToString.Length, Record.RevisionNo)
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          _RecordCount = -1
          Con.Open()
          Cmd.ExecuteNonQuery()
          _RecordCount = Cmd.Parameters("@RowCount").Value
        End Using
      End Using
      Return _RecordCount
    End Function
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
