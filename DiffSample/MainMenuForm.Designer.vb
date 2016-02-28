<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMenuForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.execButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'execButton
        '
        Me.execButton.Location = New System.Drawing.Point(247, 114)
        Me.execButton.Name = "execButton"
        Me.execButton.Size = New System.Drawing.Size(75, 23)
        Me.execButton.TabIndex = 0
        Me.execButton.Text = "実行"
        Me.execButton.UseVisualStyleBackColor = True
        '
        'MainMenuForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(550, 308)
        Me.Controls.Add(Me.execButton)
        Me.Name = "MainMenuForm"
        Me.Text = "MainMenuForm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents execButton As Button
End Class
