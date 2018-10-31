Imports EDICommon
Public Class frmMain
  Delegate Sub ThreadedSub(ByVal lt As ListBox, ByVal lb As Label)
  Delegate Sub ThreadedShow(ByVal slzFile As String)
  Delegate Sub ThreadedNone()
  Private dsM As New EDICommon.JobFiles
  Dim WithEvents jp As JobProcess.JobProcessor = Nothing
  Private Sub cmdStart_Click(sender As Object, e As EventArgs) Handles cmdStart.Click
    cmdStart.Enabled = False
    cmdSave.Enabled = False
    cmdStart.Text = "Loading..."
    Lst1.Items.Clear()
    Dim tmp As ThreadedSub = AddressOf Start
    tmp.BeginInvoke(Lst1, lblMsg, Nothing, Nothing)
  End Sub
  Private Sub Start(ByVal lt As ListBox, ByVal lb As Label)
    jp = New JobProcess.JobProcessor(lt, lb)
    jp.jpConfig = ConfigFile.GetFile(Application.StartupPath & "\Settings.xml")
    jp.jpConfig.StartupPath = Application.StartupPath
    jp.Start()
  End Sub
  Private Sub cmdStop_Click(sender As Object, e As EventArgs) Handles cmdStop.Click
    cmdStop.Enabled = False
    cmdStop.Text = "Closing..."
    jp.StopJob()
  End Sub
  Private Sub jp_JobStarted() Handles jp.JobStarted
    If cmdStart.InvokeRequired Then
      cmdStart.Invoke(New ThreadedNone(AddressOf jp_JobStarted))
    Else
      cmdStart.Enabled = False
      cmdSave.Enabled = False
      cmdStart.Text = "Start"
      cmdStop.Enabled = True
    End If
  End Sub
  Private Sub jp_JobStopped() Handles jp.JobStopped
    If cmdStop.InvokeRequired Then
      cmdStop.Invoke(New ThreadedNone(AddressOf jp_JobStopped))
    Else
      cmdStop.Enabled = False
      cmdStop.Text = "Stop"
      cmdStart.Enabled = True
      cmdSave.Enabled = True
    End If
  End Sub

  Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
    If jp IsNot Nothing Then
      Dim wl As IO.StreamWriter = New IO.StreamWriter(jp.jpConfig.JobPath & "\" & Now.ToString("yyyyMMddHHmmss") & ".log")
      For I As Long = Lst1.Items.Count - 1 To 0 Step -1
        wl.WriteLine(Lst1.Items(I).ToString)
      Next
      wl.Close()
      wl.Dispose()
      Lst1.Items.Clear()
    End If
  End Sub

  Private Sub jp_ClearList() Handles jp.ClearList
    cmdSave_Click(Nothing, Nothing)
  End Sub

End Class
