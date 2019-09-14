﻿Imports System.Windows.Forms
Imports EDICommon
Imports iTextSharp.text
Imports iTextSharp.text.html.simpleparser
Imports iTextSharp.text.pdf
Imports System.IO
Public Class JobProcessor
  Inherits TimerSupport
  Implements IDisposable

  Public Property jpConfig As ConfigFile = Nothing
  Public Event JobStarted()
  Public Event JobStopped()
  Public Event ClearList()
  Public Event ProcessedFile(ByVal slzFile As String)
  Private cad As cadUtility = Nothing
  Private vlt As vltUtility = Nothing
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
      'Get 1 Transmittal at a time from BaaN as Job
      Dim TransmittalID As String = ""
      Dim Jobs As List(Of jobFile) = jobFile.SelectList(TransmittalID, AddressOf msg)
      Dim AllDocumentProcessed As Boolean = True
      For Each tmpJob As jobFile In Jobs
        TransmittalID = tmpJob.TransmittalID
        I += 1
        msg(I & ".====================")
        Try
          tmpJob = ProcessJobFile(tmpJob)
          RaiseEvent ProcessedFile(tmpJob.SerializedPath)
        Catch ex As Exception
          msg(ex.Message)
          AllDocumentProcessed = False
        End Try
      Next
      If TransmittalID <> "" Then
        Dim tmtlH As EDICommon.SIS.EDI.ediTmtlH = EDICommon.SIS.EDI.ediTmtlH.ediTmtlHGetByID(TransmittalID)
        If Jobs.Count <= 0 Then
          tmtlH.t_edif = 3 'Error
          EDICommon.SIS.EDI.ediTmtlH.UpdateData(tmtlH)
        ElseIf AllDocumentProcessed Then
          tmtlH.t_edif = 1 'YES
          msg("Sending E-Mail.")
          Try
            Dim mailBody As String = ""
            mailBody = SIS.EDI.ediAlerts.TmtlAlert(TransmittalID)
            msg("E-Mail sent")
            EDICommon.SIS.EDI.ediTmtlH.UpdateData(tmtlH)
            '=============================
            PublishInDMS(tmtlH, mailBody, Jobs)
            '=============================
          Catch ex As Exception
            msg("Error Sending E-Mail." & ex.Message)
          End Try
        Else
          msg("All Docs NOT be downloaded or converted to PDF: " & TransmittalID)
          tmtlH.t_edif = 3 'Error
          EDICommon.SIS.EDI.ediTmtlH.UpdateData(tmtlH)
        End If
      End If
    Catch ex As Exception
      msg(ex.Message)
    End Try
  End Sub
  Private Sub PublishInDMS(tmtlH As EDICommon.SIS.EDI.ediTmtlH, mailBody As String, jobs As List(Of jobFile))
    If tmtlH.t_issu <> "007" Then Return 'OpenKM DMS
    If tmtlH.t_type <> 4 Then Return 'Only Vendor Transmittal
    If tmtlH.t_ofbp <> "SUPI00002" Then Return 'Only YNR
    Dim dmsFolderID As Integer = tmtlH.getDMSFolderID()
    If dmsFolderID <= 0 Then Return
    Dim mainFile As String = "Transmittal_" & tmtlH.t_tran & ".PDF"
    Dim mainPath As String = jpConfig.TempFolderPath & "\" & mainFile

    Dim output = New FileStream(mainPath, FileMode.Create)
    Dim sr As New StringReader(mailBody)
    Dim pdfDoc As New iTextSharp.text.Document(PageSize.A4, 10.0F, 10.0F, 100.0F, 0.0F)
    Dim htmlparser As New HTMLWorker(pdfDoc)
    iTextSharp.text.pdf.PdfWriter.GetInstance(pdfDoc, output)
    pdfDoc.Open()
    htmlparser.Parse(sr)
    pdfDoc.Close()

    'Dim tw As System.IO.StreamWriter = New System.IO.StreamWriter(mainPath)
    'tw.Write(mailBody)
    'tw.Close()
    Dim tmpVault As SIS.EDI.ediAFile
    Dim LibORGFile As String = ""
    tmpVault = New SIS.EDI.ediAFile
    With tmpVault
      LibORGFile = SIS.EDI.ediASeries.GetNextFileName
      .t_dcid = LibORGFile
      .t_hndl = "TRANSMITTALLINES_" & ERPCompany
      .t_indx = tmtlH.t_tran
      .t_prcd = "By EDI"
      .t_fnam = mainFile
      .t_lbcd = LibraryID
      .t_atby = tmtlH.t_user
      .t_aton = tmtlH.t_date
      .t_Refcntd = 0
      .t_Refcntu = 0
    End With
    tmpVault = SIS.EDI.ediAFile.InsertData(tmpVault)
    If System.IO.File.Exists(LibraryPath & "\" & LibORGFile) Then
      System.IO.File.Delete(LibraryPath & "\" & LibORGFile)
    End If
    Try
      System.IO.File.Move(mainPath, LibraryPath & "\" & LibORGFile)
    Catch ex As Exception
    End Try
    Dim pItm As SIS.DMS.UI.apiItem = SIS.DMS.UI.GetItem(dmsFolderID)
    Dim wfs As List(Of SIS.DMS.UI.apiItem) = SIS.DMS.UI.GetAssociatedWF(dmsFolderID)
    Dim wf As SIS.DMS.UI.apiItem = Nothing
    For Each x As SIS.DMS.UI.apiItem In wfs
      If x.ConvertedStatusID = 1 Then
        wf = x
        Exit For
      End If
    Next
    If wf Is Nothing Then
      wf = New SIS.DMS.UI.apiItem
      wf.ConvertedStatusID = 1
    Else
      'Get NextWF
      wf = SIS.DMS.UI.GetTopOneChild(wf)
      If wf Is Nothing Then
        wf = New SIS.DMS.UI.apiItem
        wf.ConvertedStatusID = 1
      End If
    End If
    Dim dmsItm As New SIS.DMS.UI.apiItem
    With dmsItm
      .ItemTypeID = 3 ' File
      .Description = mainFile
      .FullDescription = mainFile
      .RevisionNo = "00"
      .StatusID = wf.ConvertedStatusID
      .CreatedBy = tmtlH.t_user
      .CreatedOn = Now
      .ParentItemID = pItm.ItemID
      .ForwardLinkedItemID = IIf(wf.ItemID = 0, "", wf.ItemID)
      .ForwardLinkedItemTypeID = IIf(wf.ItemTypeID = 0, "", wf.ItemTypeID)
      .IsgecVaultID = LibORGFile
    End With
    dmsItm = SIS.DMS.UI.apiItem.InsertData(dmsItm)
    SIS.DMS.UI.apiItem.InsertHistory(dmsItm)
    If wf.ConvertedStatusID <> 1 Then
      Dim ackItm As New SIS.DMS.UI.apiItem
      With ackItm
        .ItemTypeID = 13 'Under Approval
        .Description = dmsItm.Description
        .FullDescription = dmsItm.FullDescription
        .RevisionNo = "00"
        .StatusID = dmsItm.StatusID
        .CreatedBy = tmtlH.t_user
        .CreatedOn = Now
        .ForwardLinkedItemID = wf.ItemID
        .ForwardLinkedItemTypeID = wf.ItemTypeID
        .LinkedItemID = dmsItm.ItemID
        .LinkedItemTypeID = 3
      End With
      ackItm = SIS.DMS.UI.apiItem.InsertData(ackItm)
    End If
    For Each jFile As jobFile In jobs
      Dim jItm As New SIS.DMS.UI.apiItem
      With jItm
        .ItemTypeID = 3 ' File
        .Description = jFile.PDFFileName
        .FullDescription = jFile.PDFFileName
        .RevisionNo = jFile.RevisionNo
        .ProjectID = jFile.ProjectID
        '.KeyWords = jFile.MCDxData.Title
        '.WBSID = jFile.MCDxData.ElementID
        .StatusID = dmsItm.StatusID
        .CreatedBy = tmtlH.t_user
        .CreatedOn = Now
        .ParentItemID = dmsItm.ItemID
        .ForwardLinkedItemID = wf.ItemID
        .LinkedItemTypeID = wf.ItemTypeID
        .IsgecVaultID = jFile.t_dcid
      End With
      jItm = SIS.DMS.UI.apiItem.InsertData(jItm)
      SIS.DMS.UI.apiItem.InsertHistory(jItm)
    Next
  End Sub
  Private Function ProcessJobFile(ByVal tmpJob As jobFile) As jobFile
    'Check in IsgecVault
    Dim VaultIndex As String = tmpJob.DocumentID & "_" & tmpJob.RevisionNo
    Dim tmpVault As SIS.EDI.ediAFile = SIS.EDI.ediAFile.ediAFileGetByHandleIndex("DOCUMENTMASTERPDF_" & ERPCompany, VaultIndex)
    If tmpVault Is Nothing Then
      msg("Not Found In isgec Vault.")
      msg("Downloading from Autodesk Vault.")
      tmpJob = vlt.DownloadFile(tmpJob)
      If tmpJob.IsError Then
        msg("Error In downloading from Autodesk Vault.")
      Else
        If System.IO.Path.GetExtension(tmpJob.FileName).ToUpper <> ".DWG" Then
          EDICommon.pdfWriter.generatePDF(jpConfig.TempFolderPath & "\" & tmpJob.FileName, jpConfig.PDFFolderPath)
        Else
          tmpJob = cad.ExtractPDF(tmpJob)
        End If
        msg("PDF generated." & tmpJob.PDFFileName)
        Dim PDFFile As String = jpConfig.PDFFolderPath & "\" & tmpJob.PDFFileName
        Dim ORGFile As String = jpConfig.TempFolderPath & "\" & tmpJob.FileName
        Dim LibPDFFile As String = ""
        Dim LibORGFile As String = ""
        msg("Attaching in Isgec Vault.")
        tmpVault = New SIS.EDI.ediAFile
        With tmpVault
          LibORGFile = SIS.EDI.ediASeries.GetNextFileName
          .t_dcid = LibORGFile
          .t_hndl = "DOCUMENTMASTERORG_" & ERPCompany
          .t_indx = VaultIndex
          .t_prcd = "By EDI"
          .t_fnam = tmpJob.FileName
          .t_lbcd = LibraryID
          .t_atby = tmpJob.CreatedBy
          .t_aton = tmpJob.CreatedOn
          .t_Refcntd = 0
          .t_Refcntu = 0
        End With
        tmpVault = SIS.EDI.ediAFile.InsertData(tmpVault)
        'tmpvault for PDF file
        tmpVault = New SIS.EDI.ediAFile
        With tmpVault
          LibPDFFile = SIS.EDI.ediASeries.GetNextFileName
          .t_dcid = LibPDFFile
          .t_hndl = "DOCUMENTMASTERPDF_" & ERPCompany
          .t_indx = VaultIndex
          .t_prcd = "By EDI"
          .t_fnam = tmpJob.PDFFileName
          .t_lbcd = LibraryID
          .t_atby = tmpJob.CreatedBy
          .t_aton = tmpJob.CreatedOn
          .t_Refcntd = 0
          .t_Refcntu = 0
        End With
        tmpVault = SIS.EDI.ediAFile.InsertData(tmpVault)
        tmpJob.t_dcid = LibPDFFile
        tmpJob.t_drid = tmpVault.t_drid
        msg("Copying ORG File to Library.")
        Try
          If System.IO.File.Exists(LibraryPath & "\" & LibORGFile) Then
            System.IO.File.Delete(LibraryPath & "\" & LibORGFile)
          End If
          System.IO.File.Move(ORGFile, LibraryPath & "\" & LibORGFile)
          msg("ORG File Copied.")
        Catch ex As Exception
          msg("Error: In copying ORG File to Library.")
        End Try
        msg("Copying PDF File to Library.")
        Try
          If System.IO.File.Exists(LibraryPath & "\" & LibPDFFile) Then
            System.IO.File.Delete(LibraryPath & "\" & LibPDFFile)
          End If
          System.IO.File.Move(PDFFile, LibraryPath & "\" & LibPDFFile)
          msg("PDF File Copied.")
        Catch ex As Exception
          msg("Error: In copying PDF File to Library.")
        End Try
      End If
    End If
    If tmpVault IsNot Nothing Then
      'tmpVault will be PDF file
      'Copy Vault Handle to Transmittal Handle
      Dim tmpTr As SIS.EDI.ediAFile = SIS.EDI.ediAFile.ediAFileGetByHandleIndex("TRANSMITTALLINES_" & ERPCompany, tmpJob.AttachmentIndex)
      If tmpTr IsNot Nothing Then
        'Do Nothing
      Else
        'Copy Handle to Transmittal
        tmpTr = tmpVault.Clone
        With tmpTr
          .t_hndl = "TRANSMITTALLINES_" & ERPCompany
          .t_indx = tmpJob.AttachmentIndex
          .t_atby = tmpJob.CreatedBy
          .t_aton = tmpJob.CreatedOn
        End With
        tmpTr = SIS.EDI.ediAFile.InsertData(tmpTr)
        msg("Copied from Vault Handle")
      End If
      tmpJob.t_dcid = tmpTr.t_dcid
      tmpJob.t_dcid = tmpTr.t_dcid
    End If
    Return tmpJob
  End Function

  Public Overrides Sub Started()

    msg("Checking Configuration.")
    Try
      If Not System.IO.Directory.Exists(jpConfig.JobPathWorking) Then
        System.IO.Directory.CreateDirectory(jpConfig.JobPathWorking)
      End If
    Catch ex As Exception
    End Try

    'Dim tmpIm As Impersonator = Impersonator.Impersonate("adskvault", "192.9.200.51", "adskvault@123")

    EDICommon.DBCommon.BaaNLive = jpConfig.BaaNLive
    EDICommon.DBCommon.JoomlaLive = jpConfig.JoomlaLive
    SIS.EDI.ediAlerts.Testing = jpConfig.Testing

    Dim tmp As SIS.EDI.ediALib = SIS.EDI.ediALib.GetActiveLibrary
    LibraryPath = "\\192.9.200.146\" & tmp.t_path
    LibraryID = tmp.t_lbcd
    msg("Connecting to remote attachment library.")

    If ConnectToNetworkFunctions.connectToNetwork(LibraryPath, "X:", "administrator", "Indian@12345") Then
      msg("Remote connected.")
      RemoteLibraryConnected = True
    Else
      msg("Failed to connect Remote Library.")
    End If

    cad = New cadUtility
    cad.jp = Me
    cad.jpConfig = jpConfig
    msg("Loading CAD")
    cad.LoadCad()
    msg("Connecting Vault")
    vlt = New vltUtility(jpConfig.VaultServer, jpConfig.VaultUserName, jpConfig.VaultUserPassword)
    vlt.jp = Me
    vlt.jpConfig = jpConfig
    RaiseEvent JobStarted()
    msg("Finding Transmittal Documents")
  End Sub

  Public Overrides Sub Stopped()
    If RemoteLibraryConnected Then
      ConnectToNetworkFunctions.disconnectFromNetwork("X:")
      RemoteLibraryConnected = False
    End If
    cad.UnloadCad()
    vlt.jpConfig = Nothing
    vlt = Nothing
    RaiseEvent JobStopped()
  End Sub

  Public Shared Function IsFileAvailable(ByVal FilePath As String) As Boolean
    If Not System.IO.File.Exists(FilePath) Then Return False
    Dim fInfo As System.IO.FileInfo = Nothing
    Dim st As System.IO.FileStream = Nothing
    Try
      fInfo = New System.IO.FileInfo(FilePath)
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
      st = fInfo.Open(System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
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
