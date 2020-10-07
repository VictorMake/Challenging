<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits BaseFormKT.frmBaseKT

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.TimerUpdate = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'WindowManagerPanel1
        '
        Me.WindowManagerPanel1.ShowTitle = True
        Me.WindowManagerPanel1.Size = New System.Drawing.Size(546, 56)
        '
        'TimerUpdate
        '
        Me.TimerUpdate.Interval = 5000
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.ClientSize = New System.Drawing.Size(550, 530)
        Me.Name = "frmMain"
        Me.Tag = "Challenging"
        Me.Text = "Газогенератор"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TimerUpdate As System.Windows.Forms.Timer

End Class
