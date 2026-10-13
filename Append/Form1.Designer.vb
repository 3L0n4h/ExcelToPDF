<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button2026 = New System.Windows.Forms.Button()
        Me.Button_j026 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button2026
        '
        Me.Button2026.Location = New System.Drawing.Point(59, 38)
        Me.Button2026.Name = "Button2026"
        Me.Button2026.Size = New System.Drawing.Size(159, 80)
        Me.Button2026.TabIndex = 0
        Me.Button2026.Text = "Concat 2026"
        Me.Button2026.UseVisualStyleBackColor = True
        '
        'Button_j026
        '
        Me.Button_j026.Location = New System.Drawing.Point(59, 154)
        Me.Button_j026.Name = "Button_j026"
        Me.Button_j026.Size = New System.Drawing.Size(159, 78)
        Me.Button_j026.TabIndex = 1
        Me.Button_j026.Text = "Concat j026"
        Me.Button_j026.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.Button_j026)
        Me.Controls.Add(Me.Button2026)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button2026 As Button
    Friend WithEvents Button_j026 As Button
End Class
