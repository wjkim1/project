<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AA_Update_Form
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AA_Update_Form))
        Me.Update_List = New System.Windows.Forms.ListBox()
        Me.Updata_OK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Update_List
        '
        Me.Update_List.FormattingEnabled = True
        Me.Update_List.ItemHeight = 12
        Me.Update_List.Location = New System.Drawing.Point(12, 12)
        Me.Update_List.Name = "Update_List"
        Me.Update_List.Size = New System.Drawing.Size(573, 232)
        Me.Update_List.TabIndex = 0
        '
        'Updata_OK
        '
        Me.Updata_OK.Location = New System.Drawing.Point(486, 250)
        Me.Updata_OK.Name = "Updata_OK"
        Me.Updata_OK.Size = New System.Drawing.Size(99, 31)
        Me.Updata_OK.TabIndex = 1
        Me.Updata_OK.Text = "확인"
        Me.Updata_OK.UseVisualStyleBackColor = True
        '
        'AA_Update_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(595, 287)
        Me.Controls.Add(Me.Updata_OK)
        Me.Controls.Add(Me.Update_List)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "AA_Update_Form"
        Me.Text = "AA_Update_Form"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Update_List As ListBox
    Friend WithEvents Updata_OK As Button
End Class
