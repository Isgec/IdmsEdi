<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.sc1 = New System.Windows.Forms.SplitContainer()
    Me.cmdSave = New System.Windows.Forms.Button()
    Me.cmdStop = New System.Windows.Forms.Button()
    Me.cmdStart = New System.Windows.Forms.Button()
    Me.lblMsg = New System.Windows.Forms.Label()
    Me.Lst1 = New System.Windows.Forms.ListBox()
    CType(Me.sc1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.sc1.Panel1.SuspendLayout()
    Me.sc1.Panel2.SuspendLayout()
    Me.sc1.SuspendLayout()
    Me.SuspendLayout()
    '
    'sc1
    '
    Me.sc1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.sc1.Location = New System.Drawing.Point(0, 0)
    Me.sc1.Name = "sc1"
    Me.sc1.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'sc1.Panel1
    '
    Me.sc1.Panel1.Controls.Add(Me.cmdSave)
    Me.sc1.Panel1.Controls.Add(Me.cmdStop)
    Me.sc1.Panel1.Controls.Add(Me.cmdStart)
    Me.sc1.Panel1.Controls.Add(Me.lblMsg)
    '
    'sc1.Panel2
    '
    Me.sc1.Panel2.Controls.Add(Me.Lst1)
    Me.sc1.Size = New System.Drawing.Size(582, 267)
    Me.sc1.SplitterDistance = 47
    Me.sc1.TabIndex = 0
    '
    'cmdSave
    '
    Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSave.Location = New System.Drawing.Point(308, 13)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(100, 23)
    Me.cmdSave.TabIndex = 3
    Me.cmdSave.Text = "Save Log & Clear"
    Me.cmdSave.UseVisualStyleBackColor = True
    '
    'cmdStop
    '
    Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdStop.Enabled = False
    Me.cmdStop.Location = New System.Drawing.Point(495, 13)
    Me.cmdStop.Name = "cmdStop"
    Me.cmdStop.Size = New System.Drawing.Size(75, 23)
    Me.cmdStop.TabIndex = 2
    Me.cmdStop.Text = "Stop"
    Me.cmdStop.UseVisualStyleBackColor = True
    '
    'cmdStart
    '
    Me.cmdStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdStart.Location = New System.Drawing.Point(414, 13)
    Me.cmdStart.Name = "cmdStart"
    Me.cmdStart.Size = New System.Drawing.Size(75, 23)
    Me.cmdStart.TabIndex = 1
    Me.cmdStart.Text = "Start"
    Me.cmdStart.UseVisualStyleBackColor = True
    '
    'lblMsg
    '
    Me.lblMsg.AutoSize = True
    Me.lblMsg.Location = New System.Drawing.Point(12, 18)
    Me.lblMsg.Name = "lblMsg"
    Me.lblMsg.Size = New System.Drawing.Size(50, 13)
    Me.lblMsg.TabIndex = 0
    Me.lblMsg.Text = "Message"
    '
    'Lst1
    '
    Me.Lst1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.Lst1.FormattingEnabled = True
    Me.Lst1.Location = New System.Drawing.Point(0, 0)
    Me.Lst1.Name = "Lst1"
    Me.Lst1.Size = New System.Drawing.Size(582, 216)
    Me.Lst1.TabIndex = 0
    '
    'frmMain
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(582, 267)
    Me.Controls.Add(Me.sc1)
    Me.Name = "frmMain"
    Me.Text = "EDI-IDMS"
    Me.sc1.Panel1.ResumeLayout(False)
    Me.sc1.Panel1.PerformLayout()
    Me.sc1.Panel2.ResumeLayout(False)
    CType(Me.sc1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.sc1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

  Friend WithEvents sc1 As SplitContainer
  Friend WithEvents cmdStop As Button
  Friend WithEvents cmdStart As Button
  Friend WithEvents lblMsg As Label
  Friend WithEvents Lst1 As ListBox
  Friend WithEvents cmdSave As Button
End Class
