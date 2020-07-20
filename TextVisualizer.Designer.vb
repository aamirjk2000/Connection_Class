<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TextVisualizer
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
        Me.txtVariableValue = New System.Windows.Forms.TextBox()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.txtVariable = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtVariableValue
        '
        Me.txtVariableValue.Location = New System.Drawing.Point(12, 39)
        Me.txtVariableValue.Multiline = True
        Me.txtVariableValue.Name = "txtVariableValue"
        Me.txtVariableValue.Size = New System.Drawing.Size(364, 361)
        Me.txtVariableValue.TabIndex = 0
        Me.txtVariableValue.Text = "txtVariablevValue"
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(300, 415)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 1
        Me.btnExit.Text = "E X I T"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'txtVariable
        '
        Me.txtVariable.Location = New System.Drawing.Point(13, 13)
        Me.txtVariable.Name = "txtVariable"
        Me.txtVariable.Size = New System.Drawing.Size(363, 20)
        Me.txtVariable.TabIndex = 2
        '
        'TextVisualizer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(387, 450)
        Me.Controls.Add(Me.txtVariable)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.txtVariableValue)
        Me.Name = "TextVisualizer"
        Me.Text = "TextVisualizer"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtVariableValue As Windows.Forms.TextBox
    Friend WithEvents btnExit As Windows.Forms.Button
    Friend WithEvents txtVariable As Windows.Forms.TextBox
End Class
