Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Text
Namespace SIS.WF
  Public Class wfAlerts
    Public Shared Property Testing As Boolean = False
    Private Enum tmtlType
      Customer = 1
      Internal = 2
      Site = 3
      Vendor = 4
    End Enum
    Private Class emp
      Public Property empID As Integer = 0
      Public Property empName As String = ""
      Public Property empEMail As String = ""
      Public Property webUser As String = ""

      Public Shared Function GetEmp(ByVal empID As Integer) As emp
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " emp1.t_nama as empName,"
        mSql = mSql & " bpe1.t_mail as empEMail "
        mSql = mSql & " from ttccom001200 as emp1 "
        mSql = mSql & " left outer join tbpmdm001200 as bpe1 on emp1.t_emno=bpe1.t_emno "
        mSql = mSql & " where emp1.t_emno = '" & empID & "'"
        Dim tmp As emp = Nothing
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            tmp = New emp
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              With tmp
                .empID = empID
                If Not Convert.IsDBNull(Reader("empName")) Then .empName = Reader("empName")
                If Not Convert.IsDBNull(Reader("empEMail")) Then .empEMail = Reader("empEMail")
              End With
            End If
            Reader.Close()
          End Using
        End Using
        Return tmp
      End Function

      Public Shared Function GetReceiptCreator(ByVal tmtlID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " t_user as [user] "
        mSql = mSql & " from tdmisg134200 "
        mSql = mSql & " where t_trno = '" & tmtlID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              If Not Convert.IsDBNull(Reader("user")) Then tmp = Reader("user")
            End If
            Reader.Close()
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetPONoOfReceipt(ByVal tmtlID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_orno  "
        mSql = mSql & " from tdmisg134200 "
        mSql = mSql & " where t_trno = '" & tmtlID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
            If tmp Is Nothing Then tmp = ""
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetSiteEmailIDs(ByVal ProjectID As String, ByVal AddressID As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " t_mail  "
        mSql = mSql & " from tdmisg126200 "
        mSql = mSql & " where t_cprj = '" & ProjectID & "'"
        mSql = mSql & " and   t_cadr = '" & AddressID & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetBaaNConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
            If tmp Is Nothing Then tmp = ""
          End Using
        End Using
        Return tmp
      End Function

      Public Shared Function GetTCPOIssuer(ByVal PONo As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " IssuedBy  "
        mSql = mSql & " from pak_po "
        mSql = mSql & " where pofor='TC' and ponumber = '" & PONo & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
            If tmp Is Nothing Then tmp = ""
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetPOBuyer(ByVal PONo As String) As String
        Dim mSql As String = ""
        mSql = mSql & " select top 1 "
        mSql = mSql & " BuyerID  "
        mSql = mSql & " from pak_po "
        mSql = mSql & " where pofor='TC' and ponumber = '" & PONo & "'"
        Dim tmp As String = ""
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            Con.Open()
            tmp = Cmd.ExecuteScalar()
            If tmp Is Nothing Then tmp = ""
          End Using
        End Using
        Return tmp
      End Function
      Public Shared Function GetWebUser(ByVal webUser As String) As emp
        Dim mSql As String = ""
        mSql = mSql & " select "
        mSql = mSql & " UserFullName as empName,"
        mSql = mSql & " emailid as empEMail "
        mSql = mSql & " from aspnet_users "
        mSql = mSql & " where username = '" & webUser & "'"
        Dim tmp As emp = Nothing
        Using Con As SqlConnection = New SqlConnection(EDICommon.DBCommon.GetConnectionString())
          Using Cmd As SqlCommand = Con.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = mSql
            tmp = New emp
            Con.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader()
            If (Reader.Read()) Then
              With tmp
                .webUser = webUser
                If Not Convert.IsDBNull(Reader("empName")) Then .empName = Reader("empName")
                If Not Convert.IsDBNull(Reader("empEMail")) Then .empEMail = Reader("empEMail")
              End With
            End If
            Reader.Close()
          End Using
        End Using
        Return tmp
      End Function
    End Class
    Private Shared Function gma(ByVal tmp As emp, ByRef aErr As ArrayList, Optional ByVal State As String = "") As MailAddress
      Dim x As MailAddress = Nothing
      If tmp IsNot Nothing Then
        If tmp.empEMail <> "" Then
          Try
            x = New MailAddress(tmp.empEMail, tmp.empName)
          Catch ex As Exception
            aErr.Add(State & "=> " & ex.Message)
          End Try
        Else
          aErr.Add(State & "=> " & tmp.empID & " : " & tmp.empName)
        End If
      End If
      Return x
    End Function
    Public Shared Function Alert(ByVal dm134 As SIS.DMISG.dmisg134) As Boolean
      Dim wf As SIS.WF.wfPreOrder = SIS.WF.wfPreOrder.wfPreOrderGetByID(dm134.t_wfid)

      Dim aErr As New ArrayList
      Dim mRet As String = ""
      Dim oClient As SmtpClient = New SmtpClient("192.9.200.214", 25)
      Dim oMsg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
      oClient.Credentials = New Net.NetworkCredential("adskvaultadmin", "isgec@123")

      Dim Issuer As emp = emp.GetEmp(dm134.t_appr)
      Dim Buyer As emp = emp.GetEmp(wf.UserId)
      '===================Create Notes==============
      Dim NextNo As String = ""
      Dim tmpNote As New SIS.NT.ntNotes
      With tmpNote
        .Notes_RunningNo = 0
        .NotesId = "Notes" & NextNo
        .NotesHandle = "J_PREORDER_WORKFLOW"
        .IndexValue = dm134.t_wfid
        .Title = "ISGEC Comments"
        .Description = "Commented Document(s) are attached."
        .UserId = dm134.t_appr
        .Created_Date = Now.ToString("dd/MM/yyyy")
        .SendEmailTo = Issuer.empEMail
      End With
      tmpNote = SIS.NT.ntNotes.InsertData(tmpNote)
      'Copy Attachments

      '=============================================

      With oMsg
        .Subject = "Comments From ISGEC on Offer for: " & wf.SpecificationNo
        Dim EMailIDError As String = ""
        Dim x As MailAddress = Nothing

        x = gma(Issuer, aErr, "CS-Issuer-From")
        If x IsNot Nothing Then
          .From = x
          .CC.Add(x)
        End If

        Dim Tos() As String = wf.Supplier.Split(";".ToCharArray)
        For Each _to As String In Tos
          Try
            x = New MailAddress(_to, "")
          Catch ex As Exception
            aErr.Add("To=> " & ex.Message)
          End Try
          If x IsNot Nothing Then If Not .To.Contains(x) Then .To.Add(x)
        Next

        x = gma(Buyer, aErr, "CS-Buyer-CC")
        If x IsNot Nothing Then If Not .CC.Contains(x) Then .CC.Add(x)

        If .To.Count <= 0 Then
          x = New MailAddress("baansupport@isgec.co.in", "BaaN Support")
          If Not .To.Contains(x) Then .To.Add(x)
          x = New MailAddress("lalit@isgec.co.in", "Lalit Gupta")
        End If

        If .From Is Nothing Then
          x = New MailAddress("baansupport@isgec.co.in", "BaaN Support")
          .From = x
          If Not Testing Then
            If Not .CC.Contains(x) Then .CC.Add(x)
          End If
        End If
        '====================
        Dim TestIDs As New ArrayList
        If Testing Then
          For Each xx As MailAddress In .To
            TestIDs.Add("TO => " & xx.User & " : " & xx.Address)
          Next
          .To.Clear()
          x = New MailAddress("lalit@isgec.co.in", "Lalit Gupta")
          .To.Add(x)
          For Each xx As MailAddress In .CC
            TestIDs.Add("CC => " & xx.User & " : " & xx.Address)
          Next
          .CC.Clear()
        End If
        '====================
        .IsBodyHtml = True
        Dim tblStr As String = "" '===========Put Link here  ======EDICommon.SIS.EDI.ediTmtlH.GetHTML(TransmittalID)
        Dim Header As String = ""
        Header &= "<html xmlns=""http://www.w3.org/1999/xhtml"">"
        Header &= "<head>"
        Header &= "<title></title>"
        Header &= "<style>"
        Header &= "table{"

        Header &= "border: solid 1pt black;"
        Header &= "border-collapse:collapse;"
        Header &= "font-family: Tahoma;}"

        Header &= "td{"
        Header &= "border: solid 1pt black;"
        Header &= "font-family: Tahoma;"
        Header &= "font-size: 12px;"
        Header &= "padding: 2px 2px 4px 4px;"
        Header &= "vertical-align:middle;}"

        Header &= "a{"
        Header &= "color: blue;}"
        Header &= "a:hover{"
        Header &= "color: pink;}"

        Header &= "</style>"
        Header &= "</head>"
        Header &= "<body>"
        If aErr.Count > 0 Then
          Header &= "<br/>"
          Header &= "<br/>"
          Header &= "<table>"
          Header &= "<tr><td style=""color: red""><i><b>"
          Header &= "NOTE: Download Link could not be delivered to following recipient(s), Please update their E-Mail ID in ERP-LN."
          Header &= "</b></i></td></tr>"
          For Each Err As String In aErr
            Header &= "<tr><td color=""red""><i>"
            Header &= Err
            Header &= "</i></td></tr>"
          Next
          Header &= "</table>"
        End If
        If EMailIDError <> "" Then
          Header &= "<br/>"
          Header &= "<br/>"
          Header &= "<h3>" & EMailIDError & "</h3>"
        End If
        If Testing Then
          If TestIDs.Count > 0 Then
            Header &= "<br/>"
            Header &= "<br/>"
            Header &= "<table>"
            Header &= "<tr><td style=""color: red""><i><b>"
            Header &= "TESTING"
            Header &= "</b></i></td></tr>"
            For Each test As String In TestIDs
              Header &= "<tr><td color=""red""><i>"
              Header &= test
              Header &= "</i></td></tr>"
            Next
            Header &= "</table>"
          End If
        End If

        Header &= "<br/>"
        Header &= "<br/>"
        Header &= "<p>You are requested to submit revised Offer as per comments submitted by ISGEC."
        Header &= "</p>"
        Header &= "<br/>"
        Header &= "<p>"
        Header &= "<a href='http://cloud.isgec.co.in/SupWF/SupEnquiryResponseForm.aspx?user=" & wf.RandomNo & "'>Vendor Portal for Response Submission </a>"
        Header &= "</p>"
        Header &= "<br/>"
        Header &= "<p>You are requested to access the Vendor Portal for all communication and exchange of documents"
        Header &= "</p>"
        Header &= "<br/>"
        Header &= "<br/>"
        Header &= tblStr
        Header &= "</body></html>"
        .Body = Header
      End With
      Try
        oClient.Send(oMsg)
      Catch ex As Exception
      End Try
      Return True
    End Function

    Public Shared Function WebPOWAlert(ByVal dm134 As SIS.DMISG.dmisg134) As Boolean
      Dim wf As SIS.POW.powEnquiries = SIS.POW.powEnquiries.powEnquiriesGetByID(dm134.t_pwfd, dm134.t_wfid)

      Dim aErr As New ArrayList
      Dim mRet As String = ""
      Dim oClient As SmtpClient = New SmtpClient("192.9.200.214", 25)
      Dim oMsg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
      oClient.Credentials = New Net.NetworkCredential("adskvaultadmin", "isgec@123")

      Dim Issuer As emp = emp.GetEmp(dm134.t_appr)
      Dim Buyer As emp = emp.GetEmp(wf.CreatedBy)

      With oMsg
        .Subject = "Comments From ISGEC on Offer for Enquiry: " & wf.EnquiryID
        Dim EMailIDError As String = ""
        Dim x As MailAddress = Nothing

        x = gma(Issuer, aErr, "CS-Issuer-From")
        If x IsNot Nothing Then
          .From = x
          .CC.Add(x)
        End If

        Dim Tos() As String = wf.SupplierEMailID.Split(",;".ToCharArray)
        For Each _to As String In Tos
          Try
            x = New MailAddress(_to.Trim, "")
          Catch ex As Exception
            aErr.Add("To=> " & ex.Message)
          End Try
          If x IsNot Nothing Then If Not .To.Contains(x) Then .To.Add(x)
        Next

        Tos = wf.AdditionalEMailIDs.Split(",;".ToCharArray)
        For Each _to As String In Tos
          Try
            x = New MailAddress(_to.Trim, "")
          Catch ex As Exception
            aErr.Add("To=> " & ex.Message)
          End Try
          If x IsNot Nothing Then If Not .CC.Contains(x) Then .CC.Add(x)
        Next

        x = gma(Buyer, aErr, "CS-Buyer-CC")
        If x IsNot Nothing Then If Not .CC.Contains(x) Then .CC.Add(x)

        If .To.Count <= 0 Then
          x = New MailAddress("baansupport@isgec.co.in", "BaaN Support")
          If Not .To.Contains(x) Then .To.Add(x)
          x = New MailAddress("lalit@isgec.co.in", "Lalit Gupta")
        End If

        If .From Is Nothing Then
          x = New MailAddress("baansupport@isgec.co.in", "BaaN Support")
          .From = x
          If Not Testing Then
            If Not .CC.Contains(x) Then .CC.Add(x)
          End If
        End If
        '====================
        Dim TestIDs As New ArrayList
        If Testing Then
          For Each xx As MailAddress In .To
            TestIDs.Add("TO => " & xx.User & " : " & xx.Address)
          Next
          .To.Clear()
          x = New MailAddress("lalit@isgec.co.in", "Lalit Gupta")
          .To.Add(x)
          For Each xx As MailAddress In .CC
            TestIDs.Add("CC => " & xx.User & " : " & xx.Address)
          Next
          .CC.Clear()
        End If
        '====================
        .IsBodyHtml = True
        Dim tblStr As String = "" '===========Put Link here  ======EDICommon.SIS.EDI.ediTmtlH.GetHTML(TransmittalID)
        Dim Header As String = ""
        Header &= "<html xmlns=""http://www.w3.org/1999/xhtml"">"
        Header &= "<head>"
        Header &= "<title></title>"
        Header &= "<style>"
        Header &= "table{"

        Header &= "border: solid 1pt black;"
        Header &= "border-collapse:collapse;"
        Header &= "font-family: Tahoma;}"

        Header &= "td{"
        Header &= "border: solid 1pt black;"
        Header &= "font-family: Tahoma;"
        Header &= "font-size: 12px;"
        Header &= "padding: 2px 2px 4px 4px;"
        Header &= "vertical-align:middle;}"

        Header &= "a{"
        Header &= "color: blue;}"
        Header &= "a:hover{"
        Header &= "color: pink;}"

        Header &= "</style>"
        Header &= "</head>"
        Header &= "<body>"
        'If aErr.Count > 0 Then
        '  Header &= "<br/>"
        '  Header &= "<br/>"
        '  Header &= "<table>"
        '  Header &= "<tr><td style=""color: red""><i><b>"
        '  Header &= "NOTE: Download Link could not be delivered to following recipient(s), Please update their E-Mail ID in ERP-LN."
        '  Header &= "</b></i></td></tr>"
        '  For Each Err As String In aErr
        '    Header &= "<tr><td color=""red""><i>"
        '    Header &= Err
        '    Header &= "</i></td></tr>"
        '  Next
        '  Header &= "</table>"
        'End If
        'If EMailIDError <> "" Then
        '  Header &= "<br/>"
        '  Header &= "<br/>"
        '  Header &= "<h3>" & EMailIDError & "</h3>"
        'End If
        If Testing Then
          If TestIDs.Count > 0 Then
            Header &= "<br/>"
            Header &= "<br/>"
            Header &= "<table>"
            Header &= "<tr><td style=""color: red""><i><b>"
            Header &= "TESTING"
            Header &= "</b></i></td></tr>"
            For Each test As String In TestIDs
              Header &= "<tr><td color=""red""><i>"
              Header &= test
              Header &= "</i></td></tr>"
            Next
            Header &= "</table>"
          End If
        End If
        tblStr = GetEnquiryHTML(wf)
        Header &= "<br/>"
        Header &= "<br/>"
        Header &= "<p>You are requested to submit revised Offer as per comments submitted by ISGEC."
        Header &= "</p>"
        Header &= "<br/>"
        Header &= "<p>"
        Header &= "<a href='http://cloud.isgec.co.in/WebPOW1/bsLogin.aspx?EnqKey=" & wf.VendorKey & "'>Vendor Portal for Response Submission </a>"
        Header &= "</p>"
        Header &= "<br/>"
        Header &= "<p>You are requested to access the Vendor Portal for all communication and exchange of documents"
        Header &= "</p>"
        Header &= "<br/>"
        Header &= "<br/>"
        Header &= tblStr
        Header &= "</body></html>"
        .Body = Header
      End With
      Try
        oClient.Send(oMsg)
      Catch ex As Exception
      End Try
      Return True
    End Function
    Public Shared Function GetEnquiryHTML(enq As SIS.POW.powEnquiries) As String
      Dim mRet As String = ""
      mRet &= "<table style='border:solid 1pt darkgray;'>"
      mRet &= "<tr>"
      mRet &= "<td style='padding:4px;'><b>Supplier: "
      mRet &= "</b></td>"
      mRet &= "<td style='padding:4px;'>" & IIf(enq.SupplierID <> "", enq.VR_BusinessPartner3_BPName, enq.SupplierName)
      mRet &= "</td>"
      mRet &= "</tr>"
      mRet &= "<tr>"
      mRet &= "<td style='padding:4px;'><b>Item Description: "
      mRet &= "</b></td>"
      mRet &= "<td style='padding:4px;'>" & enq.EMailSubject
      mRet &= "</td>"
      mRet &= "</tr>"
      mRet &= "<tr>"
      mRet &= "<td style='padding:4px;'><b>Remarks: "
      mRet &= "</b></td>"
      mRet &= "<td style='padding:4px;'>" & enq.EMailBody
      mRet &= "</td>"
      mRet &= "</tr>"
      mRet &= "</table>"
      Return mRet
    End Function
  End Class
End Namespace
