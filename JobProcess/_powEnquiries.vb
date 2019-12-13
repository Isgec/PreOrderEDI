Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Namespace SIS.POW
  <DataObject()> _
  Partial Public Class powEnquiries
    Private _CreatedBy As String = ""
    Private _CreatedOn As String = ""
    Private _aspnet_Users5_UserFullName As String = ""
    Private _EnquiryID As Int32 = 0
    Private _SupplierID As String = ""
    Private _SupplierName As String = ""
    Private _EMailSubject As String = ""
    Private _StatusID As String = ""
    Private _SentOn As String = ""
    Private _TSID As Int32 = 0
    Private _AdditionalEMailIDs As String = ""
    Private _EMailBody As String = ""
    Private _SupplierLoginID As String = ""
    Private _VendorKey As String = ""
    Private _SupplierEMailID As String = ""
    Private _POW_EnquiryStates1_Description As String = ""
    Private _POW_TechnicalSpecifications2_TSDescription As String = ""
    Private _VR_BusinessPartner3_BPName As String = ""
    Private _aspnet_Users4_UserFullName As String = ""
    Public Property CommercialNegotiationCompletedOn As String = ""
    Public Property SupplierFromEMailID As String = ""
    Public Property OfferReceivedOn As String = ""
    Public Property CreatedBy() As String
      Get
        Return _CreatedBy
      End Get
      Set(ByVal value As String)
        If Convert.IsDBNull(value) Then
          _CreatedBy = ""
        Else
          _CreatedBy = value
        End If
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
    Public Property aspnet_Users5_UserFullName() As String
      Get
        Return _aspnet_Users5_UserFullName
      End Get
      Set(ByVal value As String)
        _aspnet_Users5_UserFullName = value
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
    Public Property SupplierID() As String
      Get
        Return _SupplierID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SupplierID = ""
         Else
           _SupplierID = value
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
    Public Property SentOn() As String
      Get
        If Not _SentOn = String.Empty Then
          Return Convert.ToDateTime(_SentOn).ToString("dd/MM/yyyy HH:mm")
        End If
        Return _SentOn
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SentOn = ""
         Else
           _SentOn = value
         End If
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
    Public Property AdditionalEMailIDs() As String
      Get
        Return _AdditionalEMailIDs
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _AdditionalEMailIDs = ""
         Else
           _AdditionalEMailIDs = value
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
    Public Property SupplierLoginID() As String
      Get
        Return _SupplierLoginID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SupplierLoginID = ""
         Else
           _SupplierLoginID = value
         End If
      End Set
    End Property
    Public Property VendorKey() As String
      Get
        Return _VendorKey
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _VendorKey = ""
         Else
           _VendorKey = value
         End If
      End Set
    End Property
    Public Property SupplierEMailID() As String
      Get
        Return _SupplierEMailID
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _SupplierEMailID = ""
         Else
           _SupplierEMailID = value
         End If
      End Set
    End Property
    Public Property POW_EnquiryStates1_Description() As String
      Get
        Return _POW_EnquiryStates1_Description
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _POW_EnquiryStates1_Description = ""
         Else
           _POW_EnquiryStates1_Description = value
         End If
      End Set
    End Property
    Public Property POW_TechnicalSpecifications2_TSDescription() As String
      Get
        Return _POW_TechnicalSpecifications2_TSDescription
      End Get
      Set(ByVal value As String)
         If Convert.IsDBNull(Value) Then
           _POW_TechnicalSpecifications2_TSDescription = ""
         Else
           _POW_TechnicalSpecifications2_TSDescription = value
         End If
      End Set
    End Property
    Public Property VR_BusinessPartner3_BPName() As String
      Get
        Return _VR_BusinessPartner3_BPName
      End Get
      Set(ByVal value As String)
        _VR_BusinessPartner3_BPName = value
      End Set
    End Property
    Public Property aspnet_Users4_UserFullName() As String
      Get
        Return _aspnet_Users4_UserFullName
      End Get
      Set(ByVal value As String)
        _aspnet_Users4_UserFullName = value
      End Set
    End Property
    Public Readonly Property DisplayField() As String
      Get
        Return "" & _EMailSubject.ToString.PadRight(100, " ")
      End Get
    End Property
    Public Readonly Property PrimaryKey() As String
      Get
        Return _TSID & "|" & _EnquiryID
      End Get
    End Property
    Public Class PKpowEnquiries
      Private _TSID As Int32 = 0
      Private _EnquiryID As Int32 = 0
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
    End Class
    Public Shared Function powEnquiriesGetByID(ByVal TSID As Int32, ByVal EnquiryID As Int32) As SIS.POW.powEnquiries
      Dim Results As SIS.POW.powEnquiries = Nothing
      Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
        Using Cmd As SqlCommand = Con.CreateCommand()
          Cmd.CommandType = CommandType.StoredProcedure
          Cmd.CommandText = "sppowEnquiriesSelectByID"
          EDICommon.DBCommon.AddDBParameter(Cmd, "@TSID", SqlDbType.Int, TSID.ToString.Length, TSID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@EnquiryID", SqlDbType.Int, EnquiryID.ToString.Length, EnquiryID)
          EDICommon.DBCommon.AddDBParameter(Cmd, "@LoginID", SqlDbType.NVarChar, 9, "")
          Con.Open()
          Dim Reader As SqlDataReader = Cmd.ExecuteReader()
          If Reader.Read() Then
            Results = New SIS.POW.powEnquiries(Reader)
          End If
          Reader.Close()
        End Using
      End Using
      Return Results
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
