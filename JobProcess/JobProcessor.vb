Imports System.Windows.Forms
Imports EDICommon
Imports ejiVault
Public Class JobProcessor
  Inherits TimerSupport
  Implements IDisposable

  Public Property jpConfig As ConfigFile = Nothing
  Public Event JobStarted()
  Public Event JobStopped()
  Public Event ClearList()
  Public Event ProcessedFile(ByVal slzFile As String)
  Private lst As ListBox = Nothing
  Private lbl As Label = Nothing
  Private LibraryPath As String = ""
  Private LibraryID As String = ""
  Delegate Sub showMsg(ByVal str As String)
  Private RemoteLibraryConnected As Boolean = False
  Private ERPCompany As String = "200"

  Public Sub msg(ByVal str As String)
    If lbl.InvokeRequired Then
      lbl.Invoke(New showMsg(AddressOf sMsg), str)
    End If
  End Sub
  Public Sub sMsg(ByVal str As String)
    If lst.Items.Count > 4999 Then
      RaiseEvent ClearList()
      Dim I As Integer = 0
      Do While lst.Items.Count > 0
        I += 1
        Threading.Thread.Sleep(1000)
        If I > 15 Then
          Exit Do
        End If
      Loop
    End If
    lst.Items.Insert(0, str)
    lbl.Text = str
  End Sub
  Public Overrides Sub Process()
    Try
      Dim I As Integer = 0
      Dim csReceipts As List(Of SIS.DMISG.dmisg134) = SIS.DMISG.dmisg134.CommentSubmittedReceipts(jpConfig.SinceDays)
      msg("Total Receipts: " & SIS.DMISG.dmisg134.RecordCount)
      For Each dm134 As SIS.DMISG.dmisg134 In csReceipts
        I += 1
        msg(I & "=> " & dm134.t_rcno)
        If dm134.t_wfid >= 10000 Then
          If dm134.t_stat = 5 Then
            'Technically cleared
            Dim tmp As SIS.POW.powOffers = SIS.POW.powOffers.powOffersGetByReceiptRevision(dm134.t_rcno, dm134.t_revn)
            If tmp IsNot Nothing Then
              With tmp
                .ERPStatusID = 5 'Technically Cleared
                .EvaluatedBy = dm134.t_appr
                Try
                  .EValuatedOn = dm134.t_adat
                Catch ex As Exception
                  .EValuatedOn = Now
                End Try
              End With
              tmp = SIS.POW.powOffers.UpdateData(tmp)
            End If
          Else
            'Comment Submitted
            Dim tmp As SIS.ERP.erpIDMSAlertSent = SIS.ERP.erpIDMSAlertSent.erpIDMSAlertSentGetByID(dm134.t_rcno, dm134.t_revn)
            If tmp IsNot Nothing Then
              msg(" ====> alert already sent.")
              Continue For
            End If
            Try
              msg("Sending WebPOW Alert: " & dm134.t_rcno)
              Dim Record As New SIS.POW.powOffers
              With Record
                .RecordTypeID = 5 'communication
                .RecordRevision = "00"
                .EMailSubject = "ISGEC Comments"
                .SubmittedBy = dm134.t_user
                .CreatedOn = Now
                .StatusID = 6 'Comment, 7=>Cleared
                .EMailBody = "Pl. find attached documents."
                .EnquiryID = dm134.t_wfid
                .TSID = dm134.t_pwfd
                .ReceiptID = dm134.t_rcno
                .ReceiptRevision = dm134.t_revn
                .EvaluatedBy = dm134.t_appr
                Try
                  .EValuatedOn = dm134.t_adat
                Catch ex As Exception
                  .EValuatedOn = Now
                End Try
                .SubmittedByBuyer = True
                .ForSupplier = False
              End With
              Record = SIS.POW.powOffers.InsertData(Record)
              'Copy Commented Attachments
              Dim comments As List(Of EJI.ediAFile) = EJI.ediAFile.GetFilesByHandleLikeIndex("IDMSEVALUATOR_200", dm134.t_rcno & "_" & dm134.t_revn & "_")
              For Each cmt As EJI.ediAFile In comments
                cmt.t_hndl = Record.AthHandle
                cmt.t_indx = Record.AthIndex
                EJI.ediAFile.InsertData(cmt)
              Next
              If SIS.WF.wfAlerts.WebPOWAlert(dm134) Then
                tmp = New SIS.ERP.erpIDMSAlertSent
                tmp.ReceiptNo = dm134.t_rcno
                tmp.RevisionNo = dm134.t_revn
                tmp.MailSentOn = Now
                SIS.ERP.erpIDMSAlertSent.InsertData(tmp)
                msg("Alert Sent: " & dm134.t_rcno)
              End If
            Catch ex As Exception
              msg(ex.Message)
            End Try
          End If
        Else
          If dm134.t_stat = 4 Then
            'comment Submitted
            Dim tmp As SIS.ERP.erpIDMSAlertSent = SIS.ERP.erpIDMSAlertSent.erpIDMSAlertSentGetByID(dm134.t_rcno, dm134.t_revn)
            If tmp IsNot Nothing Then
              msg(" ====> alert already sent.")
              Continue For
            End If
            Try
              msg("Sending Alert: " & dm134.t_rcno)
              If SIS.WF.wfAlerts.Alert(dm134) Then
                tmp = New SIS.ERP.erpIDMSAlertSent
                tmp.ReceiptNo = dm134.t_rcno
                tmp.RevisionNo = dm134.t_revn
                tmp.MailSentOn = Now
                SIS.ERP.erpIDMSAlertSent.InsertData(tmp)
                msg("Alert Sent: " & dm134.t_rcno)
              End If
            Catch ex As Exception
              msg(ex.Message)
            End Try
          Else
            'Technically Cleared
            'Nothing to be done as receipt is not In Joomla
            'Though Trying to get record in Joomla, which will be "Nothing"
            Dim tmp As SIS.POW.powOffers = SIS.POW.powOffers.powOffersGetByReceiptRevision(dm134.t_rcno, dm134.t_revn)
            If tmp IsNot Nothing Then
              With tmp
                .ERPStatusID = 5 'Technically Cleared
                .EvaluatedBy = dm134.t_appr
                Try
                  .EValuatedOn = dm134.t_adat
                Catch ex As Exception
                  .EValuatedOn = Now
                End Try
              End With
              tmp = SIS.POW.powOffers.UpdateData(tmp)
            End If
          End If
        End If
        If IsStopping Then
          Exit For
        End If
      Next
    Catch ex As Exception
      msg(ex.Message)
    End Try
  End Sub
  Public Overrides Sub Started()

    msg("Checking Configuration.")
    Try
      If Not IO.Directory.Exists(jpConfig.JobPathWorking) Then
        IO.Directory.CreateDirectory(jpConfig.JobPathWorking)
      End If
    Catch ex As Exception
    End Try


    EDICommon.DBCommon.BaaNLive = jpConfig.BaaNLive
    EDICommon.DBCommon.JoomlaLive = jpConfig.JoomlaLive
    SIS.WF.wfAlerts.Testing = jpConfig.Testing
    EJI.DBCommon.BaaNLive = jpConfig.BaaNLive
    EJI.DBCommon.JoomlaLive = jpConfig.JoomlaLive
    EJI.DBCommon.ERPCompany = jpConfig.ISGECVaultCompany
    EJI.DBCommon.IsLocalISGECVault = jpConfig.IsLocalISGECVault
    EJI.DBCommon.ISGECVaultIP = jpConfig.ISGECVaultIP


    Dim tmp As EJI.ediALib = EJI.ediALib.GetActiveLibrary
    LibraryPath = tmp.LibraryPath
    LibraryID = tmp.t_lbcd
    If Not jpConfig.IsLocalISGECVault Then
      msg("Connecting to remote attachment library.")
      If EJI.ediALib.ConnectISGECVault(tmp) Then
        msg("Remote connected.")
        RemoteLibraryConnected = True
      Else
        msg("Failed to connect Remote Library.")
      End If
    End If

    RaiseEvent JobStarted()
    msg("Finding Transmittal Documents")
  End Sub
  Public Overrides Sub Stopped()
    If RemoteLibraryConnected Then
      ConnectToNetworkFunctions.disconnectFromNetwork("X:")
      RemoteLibraryConnected = False
    End If
    RaiseEvent JobStopped()
  End Sub

  Public Shared Function IsFileAvailable(ByVal FilePath As String) As Boolean
    If Not IO.File.Exists(FilePath) Then Return False
    Dim fInfo As IO.FileInfo = Nothing
    Dim st As IO.FileStream = Nothing
    Try
      fInfo = New IO.FileInfo(FilePath)
    Catch ex As Exception
      Return False
    End Try
    Dim ret As Boolean = False
    If fInfo.IsReadOnly Then
      If DateDiff(DateInterval.Minute, fInfo.CreationTime, Now) >= 1 Then
        fInfo.IsReadOnly = False
      End If
    End If
    Try
      st = fInfo.Open(IO.FileMode.Open, IO.FileAccess.ReadWrite, IO.FileShare.None)
      ret = True
    Catch ex As Exception
      ret = False
    Finally
      If st IsNot Nothing Then
        st.Close()
      End If
    End Try
    Return ret
  End Function
  Sub New(ByVal lt As ListBox, ByVal lb As Label)
    lst = lt
    lbl = lb
  End Sub

  Sub New()
    'dummy
  End Sub

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
        lst.Dispose()
        lbl.Dispose()
      End If

      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    ' TODO: uncomment the following line if Finalize() is overridden above.
    ' GC.SuppressFinalize(Me)
  End Sub
#End Region
End Class
