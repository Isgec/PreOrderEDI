Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.WF
  <DataObject()> _
  Partial Public Class wfPreOrder
    Private Shared _RecordCount As Integer
    Private _WFID As Int32 = 0
    Private _Parent_WFID As String = ""
    Private _Project As String = ""
    Private _Element As String = ""
    Private _SpecificationNo As String = ""
    Private _Buyer As String = ""
    Private _WF_Status As String = ""
    Private _UserId As String = ""
    Private _DateTime As String = ""
    Private _Supplier As String = ""
    Private _SupplierName As String = ""
    Private _RandomNo As String = ""
    Private _PMDLDocNo As String = ""
    Private _SupplierCode As String = ""
    Private _ReceiptNo As String = ""
    Private _Manager As String = ""
    Private _EmailSubject As String = ""
    Private _IndentNo As String = ""
    Private _IndentLine As String = ""
    Private _LotItem As String = ""
    Public Property WFID() As Int32
      Get
        Return _WFID
      End Get
      Set(ByVal value As Int32)
        _WFID = value
      End Set
    End Property
    Public Property Parent_WFID() As String
      Get
        Return _Parent_WFID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _Parent_WFID = ""
         Else
           _Parent_WFID = value
         End If
      End Set
    End Property
    Public Property Project() As String
      Get
        Return _Project
      End Get
      Set(ByVal value As String)
        _Project = value
      End Set
    End Property
    Public Property Element() As String
      Get
        Return _Element
      End Get
      Set(ByVal value As String)
        _Element = value
      End Set
    End Property
    Public Property SpecificationNo() As String
      Get
        Return _SpecificationNo
      End Get
      Set(ByVal value As String)
        _SpecificationNo = value
      End Set
    End Property
    Public Property Buyer() As String
      Get
        Return _Buyer
      End Get
      Set(ByVal value As String)
        _Buyer = value
      End Set
    End Property
    Public Property WF_Status() As String
      Get
        Return _WF_Status
      End Get
      Set(ByVal value As String)
        _WF_Status = value
      End Set
    End Property
    Public Property UserId() As String
      Get
        Return _UserId
      End Get
      Set(ByVal value As String)
        _UserId = value
      End Set
    End Property
    Public Property DateTime() As String
      Get
        If Not _DateTime = String.Empty Then
          Return Convert.ToDateTime(_DateTime).ToString("dd/MM/yyyy")
        End If
        Return _DateTime
      End Get
      Set(ByVal value As String)
         _DateTime = value
      End Set
    End Property
    Public Property Supplier() As String
      Get
        Return _Supplier
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _Supplier = ""
         Else
           _Supplier = value
         End If
      End Set
    End Property
    Public Property SupplierName() As String
      Get
        Return _SupplierName
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SupplierName = ""
         Else
           _SupplierName = value
         End If
      End Set
    End Property
    Public Property RandomNo() As String
      Get
        Return _RandomNo
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _RandomNo = ""
         Else
           _RandomNo = value
         End If
      End Set
    End Property
    Public Property PMDLDocNo() As String
      Get
        Return _PMDLDocNo
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _PMDLDocNo = ""
         Else
           _PMDLDocNo = value
         End If
      End Set
    End Property
    Public Property SupplierCode() As String
      Get
        Return _SupplierCode
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SupplierCode = ""
         Else
           _SupplierCode = value
         End If
      End Set
    End Property
    Public Property ReceiptNo() As String
      Get
        Return _ReceiptNo
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ReceiptNo = ""
         Else
           _ReceiptNo = value
         End If
      End Set
    End Property
    Public Property Manager() As String
      Get
        Return _Manager
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _Manager = ""
         Else
           _Manager = value
         End If
      End Set
    End Property
    Public Property EmailSubject() As String
      Get
        Return _EmailSubject
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _EmailSubject = ""
         Else
           _EmailSubject = value
         End If
      End Set
    End Property
    Public Property IndentNo() As String
      Get
        Return _IndentNo
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _IndentNo = ""
         Else
           _IndentNo = value
         End If
      End Set
    End Property
    Public Property IndentLine() As String
      Get
        Return _IndentLine
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _IndentLine = ""
         Else
           _IndentLine = value
         End If
      End Set
    End Property
    Public Property LotItem() As String
      Get
        Return _LotItem
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _LotItem = ""
         Else
           _LotItem = value
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
        Return _WFID
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
    Public Class PKwfPreOrder
      Private _WFID As Int32 = 0
      Public Property WFID() As Int32
        Get
          Return _WFID
        End Get
        Set(ByVal value As Int32)
          _WFID = value
        End Set
      End Property
    End Class
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function wfPreOrderGetNewRecord() As SIS.WF.wfPreOrder
      Return New SIS.WF.wfPreOrder()
    End Function
    <DataObjectMethod(DataObjectMethodType.Select)> _
    Public Shared Function wfPreOrderGetByID(ByVal WFID As Int32) As SIS.WF.wfPreOrder
      Dim Results As SIS.WF.wfPreOrder = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.Text
          Cmd.CommandText = "select * from WF1_PreOrder where WFID=" & WFID
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.WF.wfPreOrder(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    '<DataObjectMethod(DataObjectMethodType.Select)>
    'Public Shared Function wfPreOrderSelectList(ByVal StartRowIndex As Integer, ByVal MaximumRows As Integer, ByVal OrderBy As String, ByVal SearchState As Boolean, ByVal SearchText As String) As List(Of SIS.WF.wfPreOrder)
    '  Dim Results As List(Of SIS.WF.wfPreOrder) = Nothing
    '  Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
    '    Using Cmd As SqlCommand = Con.CreateCommand()
    '      Cmd.CommandType = CommandType.StoredProcedure
    '      If SearchState Then
    '        Cmd.CommandText = "spwfPreOrderSelectListSearch"
    '        EDICommon.DBCommon.AddDBParameter(Cmd, "@KeyWord", SqlDbType.NVarChar, 250, SearchText)
    '      Else
    '        Cmd.CommandText = "spwfPreOrderSelectListFilteres"
    '      End If
    '      EDICommon.DBCommon.AddDBParameter(Cmd, "@StartRowIndex", SqlDbType.Int, -1, StartRowIndex)
    '      EDICommon.DBCommon.AddDBParameter(Cmd, "@MaximumRows", SqlDbType.Int, -1, MaximumRows)
    '      EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
    '      EDICommon.DBCommon.AddDBParameter(Cmd, "@OrderBy", SqlDbType.NVarChar, 50, OrderBy)
    '      Cmd.Parameters.Add("@RecordCount", SqlDbType.Int)
    '      Cmd.Parameters("@RecordCount").Direction = ParameterDirection.Output
    '      _RecordCount = -1
    '      Results = New List(Of SIS.WF.wfPreOrder)()
    '      Con.Open()
    '      Dim Reader As SqlDataReader = Cmd.ExecuteReader()
    '      While (Reader.Read())
    '        Results.Add(New SIS.WF.wfPreOrder(Reader))
    '      End While
    '      Reader.Close()
    '      _RecordCount = Cmd.Parameters("@RecordCount").Value
    '    End Using
    '  End Using
    '  Return Results
    'End Function
    Public Shared Function wfPreOrderSelectCount(ByVal SearchState As Boolean, ByVal SearchText As String) As Integer
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
