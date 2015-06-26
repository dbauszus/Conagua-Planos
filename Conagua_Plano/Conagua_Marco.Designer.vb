<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Conagua_Marco
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Conagua_Marco))
        Me.cbEscala = New System.Windows.Forms.ComboBox()
        Me.btnCrearMarco = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cbEscala
        '
        Me.cbEscala.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbEscala.FormattingEnabled = True
        Me.cbEscala.Items.AddRange(New Object() {"500", "1000", "2000", "5000"})
        Me.cbEscala.Location = New System.Drawing.Point(12, 12)
        Me.cbEscala.Name = "cbEscala"
        Me.cbEscala.Size = New System.Drawing.Size(258, 24)
        Me.cbEscala.TabIndex = 0
        '
        'btnCrearMarco
        '
        Me.btnCrearMarco.Location = New System.Drawing.Point(12, 42)
        Me.btnCrearMarco.Name = "btnCrearMarco"
        Me.btnCrearMarco.Size = New System.Drawing.Size(258, 23)
        Me.btnCrearMarco.TabIndex = 1
        Me.btnCrearMarco.Text = "Crear Marco"
        Me.btnCrearMarco.UseVisualStyleBackColor = True
        '
        'Conagua_Marco
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(282, 79)
        Me.Controls.Add(Me.btnCrearMarco)
        Me.Controls.Add(Me.cbEscala)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Conagua_Marco"
        Me.Text = "Marco"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cbEscala As System.Windows.Forms.ComboBox
    Friend WithEvents btnCrearMarco As System.Windows.Forms.Button
End Class
