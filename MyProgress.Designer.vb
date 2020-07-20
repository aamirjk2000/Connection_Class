<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MyProgress
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
        Me.MyProgressBar = New System.Windows.Forms.ProgressBar()
        Me.MyLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'MyProgressBar
        '
        Me.MyProgressBar.Location = New System.Drawing.Point(12, 34)
        Me.MyProgressBar.Name = "MyProgressBar"
        Me.MyProgressBar.Size = New System.Drawing.Size(366, 23)
        Me.MyProgressBar.TabIndex = 0
        '
        'MyLabel
        '
        Me.MyLabel.AutoSize = True
        Me.MyLabel.Location = New System.Drawing.Point(13, 15)
        Me.MyLabel.Name = "MyLabel"
        Me.MyLabel.Size = New System.Drawing.Size(50, 13)
        Me.MyLabel.TabIndex = 1
        Me.MyLabel.Text = "Message"
        '
        'MyProgress
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(392, 76)
        Me.Controls.Add(Me.MyLabel)
        Me.Controls.Add(Me.MyProgressBar)
        Me.Name = "MyProgress"
        Me.Text = "MyProgress"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MyProgressBar As Windows.Forms.ProgressBar
    Public WithEvents MyLabel As Windows.Forms.Label
End Class
