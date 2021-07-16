Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.POW
  <DataObject()> _
  Partial Public Class powOffers
    Private _ForSupplier As Boolean = False
    Private _CreatedOn As String = ""
    Private _RecordID As Int32 = 0
    Private _RecordTypeID As String = ""
    Private _RecordRevision As String = ""
    Private _EMailSubject As String = ""
    Private _SubmittedBy As String = ""
    Private _SubmittedOn As String = ""
    Private _StatusID As String = ""
    Private _EMailBody As String = ""
    Private _EvaluatedBy As String = ""
    Private _EnquiryID As Int32 = 0
    Private _TSID As Int32 = 0
    Private _DistributedOn As String = ""
    Private _ReceiptID As String = ""
    Private _ReceiptRevision As String = ""
    Private _SubmittedByBuyer As Boolean = False
    Private _EValuatedOn As String = ""
    Private _AcknowledgedOn As String = ""
    Private _aspnet_Users1_UserFullName As String = ""
    Private _aspnet_Users2_UserFullName As String = ""
    Private _POW_Enquiries3_EMailSubject As String = ""
    Private _POW_OfferStates4_Description As String = ""
    Private _POW_RecordTypes5_Description As String = ""
    Private _POW_TechnicalSpecifications6_TSDescription As String = ""
    Public Property ERPStatusID As String = ""
    Public Property ForSupplier() As Boolean
      Get
        Return _ForSupplier
      End Get
      Set(ByVal value As Boolean)
        _ForSupplier = value
      End Set
    End Property
    Public Property CreatedOn() As String
      Get
        If Not _CreatedOn = String.Empty Then
          Return Convert.ToDateTime(_CreatedOn).ToString("dd/MM/yyyy HH:mm")
        End If
        Return _CreatedOn
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _CreatedOn = ""
        Else
          _CreatedOn = value
        End If
      End Set
    End Property
    Public Property RecordID() As Int32
      Get
        Return _RecordID
      End Get
      Set(ByVal value As Int32)
        _RecordID = value
      End Set
    End Property
    Public Property RecordTypeID() As String
      Get
        Return _RecordTypeID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _RecordTypeID = ""
         Else
           _RecordTypeID = value
         End If
      End Set
    End Property
    Public Property RecordRevision() As String
      Get
        Return _RecordRevision
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _RecordRevision = ""
         Else
           _RecordRevision = value
         End If
      End Set
    End Property
    Public Property EMailSubject() As String
      Get
        Return _EMailSubject
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _EMailSubject = ""
         Else
           _EMailSubject = value
         End If
      End Set
    End Property
    Public Property SubmittedBy() As String
      Get
        Return _SubmittedBy
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SubmittedBy = ""
         Else
           _SubmittedBy = value
         End If
      End Set
    End Property
    Public Property SubmittedOn() As String
      Get
        If Not _SubmittedOn = String.Empty Then
          Return Convert.ToDateTime(_SubmittedOn).ToString("dd/MM/yyyy HH:mm")
        End If
        Return _SubmittedOn
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SubmittedOn = ""
         Else
           _SubmittedOn = value
         End If
      End Set
    End Property
    Public Property StatusID() As String
      Get
        Return _StatusID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _StatusID = ""
         Else
           _StatusID = value
         End If
      End Set
    End Property
    Public Property EMailBody() As String
      Get
        Return _EMailBody
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _EMailBody = ""
         Else
           _EMailBody = value
         End If
      End Set
    End Property
    Public Property EvaluatedBy() As String
      Get
        Return _EvaluatedBy
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _EvaluatedBy = ""
         Else
           _EvaluatedBy = value
         End If
      End Set
    End Property
    Public Property EnquiryID() As Int32
      Get
        Return _EnquiryID
      End Get
      Set(ByVal value As Int32)
        _EnquiryID = value
      End Set
    End Property
    Public Property TSID() As Int32
      Get
        Return _TSID
      End Get
      Set(ByVal value As Int32)
        _TSID = value
      End Set
    End Property
    Public Property DistributedOn() As String
      Get
        If Not _DistributedOn = String.Empty Then
          Return Convert.ToDateTime(_DistributedOn).ToString("dd/MM/yyyy")
        End If
        Return _DistributedOn
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _DistributedOn = ""
         Else
           _DistributedOn = value
         End If
      End Set
    End Property
    Public Property ReceiptID() As String
      Get
        Return _ReceiptID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ReceiptID = ""
         Else
           _ReceiptID = value
         End If
      End Set
    End Property
    Public Property ReceiptRevision() As String
      Get
        Return _ReceiptRevision
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _ReceiptRevision = ""
         Else
           _ReceiptRevision = value
         End If
      End Set
    End Property
    Public Property SubmittedByBuyer() As Boolean
      Get
        Return _SubmittedByBuyer
      End Get
      Set(ByVal value As Boolean)
        _SubmittedByBuyer = value
      End Set
    End Property
    Public Property EValuatedOn() As String
      Get
        If Not _EValuatedOn = String.Empty Then
          Return Convert.ToDateTime(_EValuatedOn).ToString("dd/MM/yyyy")
        End If
        Return _EValuatedOn
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _EValuatedOn = ""
         Else
           _EValuatedOn = value
         End If
      End Set
    End Property
    Public Property AcknowledgedOn() As String
      Get
        If Not _AcknowledgedOn = String.Empty Then
          Return Convert.ToDateTime(_AcknowledgedOn).ToString("dd/MM/yyyy")
        End If
        Return _AcknowledgedOn
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _AcknowledgedOn = ""
         Else
           _AcknowledgedOn = value
         End If
      End Set
    End Property
    Public Property aspnet_Users1_UserFullName() As String
      Get
        Return _aspnet_Users1_UserFullName
      End Get
      Set(ByVal value As String)
        _aspnet_Users1_UserFullName = value
      End Set
    End Property
    Public Property aspnet_Users2_UserFullName() As String
      Get
        Return _aspnet_Users2_UserFullName
      End Get
      Set(ByVal value As String)
        _aspnet_Users2_UserFullName = value
      End Set
    End Property
    Public Property POW_Enquiries3_EMailSubject() As String
      Get
        Return _POW_Enquiries3_EMailSubject
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _POW_Enquiries3_EMailSubject = ""
         Else
           _POW_Enquiries3_EMailSubject = value
         End If
      End Set
    End Property
    Public Property POW_OfferStates4_Description() As String
      Get
        Return _POW_OfferStates4_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _POW_OfferStates4_Description = ""
         Else
           _POW_OfferStates4_Description = value
         End If
      End Set
    End Property
    Public Property POW_RecordTypes5_Description() As String
      Get
        Return _POW_RecordTypes5_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _POW_RecordTypes5_Description = ""
         Else
           _POW_RecordTypes5_Description = value
         End If
      End Set
    End Property
    Public Property POW_TechnicalSpecifications6_TSDescription() As String
      Get
        Return _POW_TechnicalSpecifications6_TSDescription
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _POW_TechnicalSpecifications6_TSDescription = ""
         Else
           _POW_TechnicalSpecifications6_TSDescription = value
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
        Return _TSID & "|" & _EnquiryID & "|" & _RecordID
      End Get
    End Property
    Public Class PKpowOffers
      Private _TSID As Int32 = 0
      Private _EnquiryID As Int32 = 0
      Private _RecordID As Int32 = 0
      Public Property TSID() As Int32
        Get
          Return _TSID
        End Get
        Set(ByVal value As Int32)
          _TSID = value
        End Set
      End Property
      Public Property EnquiryID() As Int32
        Get
          Return _EnquiryID
        End Get
        Set(ByVal value As Int32)
          _EnquiryID = value
        End Set
      End Property
      Public Property RecordID() As Int32
        Get
          Return _RecordID
        End Get
        Set(ByVal value As Int32)
          _RecordID = value
        End Set
      End Property
    End Class
    Public Property AthProcess As String = ""
    Public Property AthHandle As String = "J_PREORDER_OFFER"
    Public ReadOnly Property AthIndex As String
      Get
        Return TSID & "_" & EnquiryID & "_" & RecordID
      End Get
    End Property
    Public Shared Function powOffersGetByReceiptRevision(ByVal ReceiptID As String, ByVal ReceiptRevision As String) As SIS.POW.powOffers
      Dim Results As SIS.POW.powOffers = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppow_LG_OffersSelectByReceiptRevision"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ReceiptID", SqlDbType.NVarChar, 9, ReceiptID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ReceiptRevision", SqlDbType.NVarChar, 5, ReceiptRevision)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.POW.powOffers(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
    End Function
    Public Shared Function InsertData(ByVal Record As SIS.POW.powOffers) As SIS.POW.powOffers
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppowOffersInsert"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@RecordTypeID", SqlDbType.Int, 11, IIf(Record.RecordTypeID = "", Convert.DBNull, Record.RecordTypeID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@RecordRevision", SqlDbType.NVarChar, 6, IIf(Record.RecordRevision = "", Convert.DBNull, Record.RecordRevision))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EMailSubject", SqlDbType.NVarChar, 101, IIf(Record.EMailSubject = "", Convert.DBNull, Record.EMailSubject))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubmittedBy", SqlDbType.NVarChar, 9, IIf(Record.SubmittedBy = "", Convert.DBNull, Record.SubmittedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubmittedOn", SqlDbType.DateTime, 21, IIf(Record.SubmittedOn = "", Convert.DBNull, Record.SubmittedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StatusID", SqlDbType.Int, 11, IIf(Record.StatusID = "", Convert.DBNull, Record.StatusID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EMailBody", SqlDbType.NVarChar, 4001, IIf(Record.EMailBody = "", Convert.DBNull, Record.EMailBody))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EvaluatedBy", SqlDbType.NVarChar, 9, IIf(Record.EvaluatedBy = "", Convert.DBNull, Record.EvaluatedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EnquiryID", SqlDbType.Int, 11, Record.EnquiryID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TSID", SqlDbType.Int, 11, Record.TSID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DistributedOn", SqlDbType.DateTime, 21, IIf(Record.DistributedOn = "", Convert.DBNull, Record.DistributedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ReceiptID", SqlDbType.NVarChar, 10, IIf(Record.ReceiptID = "", Convert.DBNull, Record.ReceiptID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ReceiptRevision", SqlDbType.NVarChar, 6, IIf(Record.ReceiptRevision = "", Convert.DBNull, Record.ReceiptRevision))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubmittedByBuyer", SqlDbType.Bit, 3, Record.SubmittedByBuyer)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EValuatedOn", SqlDbType.DateTime, 21, IIf(Record.EValuatedOn = "", Convert.DBNull, Record.EValuatedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcknowledgedOn", SqlDbType.DateTime, 21, IIf(Record.AcknowledgedOn = "", Convert.DBNull, Record.AcknowledgedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ForSupplier", SqlDbType.Bit, 3, Record.ForSupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedOn", SqlDbType.DateTime, 21, IIf(Record.CreatedOn = "", Convert.DBNull, Record.CreatedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ERPStatusID", SqlDbType.Int, 11, IIf(Record.ERPStatusID = "", Convert.DBNull, Record.ERPStatusID))
          Cmd.Parameters.Add("@Return_TSID", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_TSID").Direction = ParameterDirection.Output
          Cmd.Parameters.Add("@Return_EnquiryID", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_EnquiryID").Direction = ParameterDirection.Output
          Cmd.Parameters.Add("@Return_RecordID", SqlDbType.Int, 11)
          Cmd.Parameters("@Return_RecordID").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
          Record.TSID = Cmd.Parameters("@Return_TSID").Value
          Record.EnquiryID = Cmd.Parameters("@Return_EnquiryID").Value
          Record.RecordID = Cmd.Parameters("@Return_RecordID").Value
        End Using
      End Using
      Return Record
    End Function
    Public Shared Function UpdateData(ByVal Record As SIS.POW.powOffers) As SIS.POW.powOffers
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppowOffersUpdate"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_RecordID", SqlDbType.Int, 11, Record.RecordID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_EnquiryID", SqlDbType.Int, 11, Record.EnquiryID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@Original_TSID", SqlDbType.Int, 11, Record.TSID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@RecordTypeID", SqlDbType.Int, 11, IIf(Record.RecordTypeID = "", Convert.DBNull, Record.RecordTypeID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@RecordRevision", SqlDbType.NVarChar, 6, IIf(Record.RecordRevision = "", Convert.DBNull, Record.RecordRevision))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EMailSubject", SqlDbType.NVarChar, 101, IIf(Record.EMailSubject = "", Convert.DBNull, Record.EMailSubject))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubmittedBy", SqlDbType.NVarChar, 9, IIf(Record.SubmittedBy = "", Convert.DBNull, Record.SubmittedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubmittedOn", SqlDbType.DateTime, 21, IIf(Record.SubmittedOn = "", Convert.DBNull, Record.SubmittedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@StatusID", SqlDbType.Int, 11, IIf(Record.StatusID = "", Convert.DBNull, Record.StatusID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EMailBody", SqlDbType.NVarChar, 4001, IIf(Record.EMailBody = "", Convert.DBNull, Record.EMailBody))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EvaluatedBy", SqlDbType.NVarChar, 9, IIf(Record.EvaluatedBy = "", Convert.DBNull, Record.EvaluatedBy))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EnquiryID", SqlDbType.Int, 11, Record.EnquiryID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TSID", SqlDbType.Int, 11, Record.TSID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@DistributedOn", SqlDbType.DateTime, 21, IIf(Record.DistributedOn = "", Convert.DBNull, Record.DistributedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ReceiptID", SqlDbType.NVarChar, 10, IIf(Record.ReceiptID = "", Convert.DBNull, Record.ReceiptID))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ReceiptRevision", SqlDbType.NVarChar, 6, IIf(Record.ReceiptRevision = "", Convert.DBNull, Record.ReceiptRevision))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@SubmittedByBuyer", SqlDbType.Bit, 3, Record.SubmittedByBuyer)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EValuatedOn", SqlDbType.DateTime, 21, IIf(Record.EValuatedOn = "", Convert.DBNull, Record.EValuatedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@AcknowledgedOn", SqlDbType.DateTime, 21, IIf(Record.AcknowledgedOn = "", Convert.DBNull, Record.AcknowledgedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ForSupplier", SqlDbType.Bit, 3, Record.ForSupplier)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@CreatedOn", SqlDbType.DateTime, 21, IIf(Record.CreatedOn = "", Convert.DBNull, Record.CreatedOn))
          EDICommon.DBCommon.AddDBParameter(Cmd, "@ERPStatusID", SqlDbType.Int, 11, IIf(Record.ERPStatusID = "", Convert.DBNull, Record.ERPStatusID))
          Cmd.Parameters.Add("@RowCount", SqlDbType.Int)
          Cmd.Parameters("@RowCount").Direction = ParameterDirection.Output
          Con.Open()
          Cmd.ExecuteNonQuery()
        End Using
      End Using
      Return Record
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
