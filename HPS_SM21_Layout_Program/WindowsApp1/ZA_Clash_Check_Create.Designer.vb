<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ZA_Clash_Check_Create
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
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ZA_Clash_Check_Create))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.EPN_NO_Label = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Clash_Check_Create_Lebel = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "EPN NO :"
        '
        'EPN_NO_Label
        '
        Me.EPN_NO_Label.AutoSize = True
        Me.EPN_NO_Label.Location = New System.Drawing.Point(78, 9)
        Me.EPN_NO_Label.Name = "EPN_NO_Label"
        Me.EPN_NO_Label.Size = New System.Drawing.Size(11, 12)
        Me.EPN_NO_Label.TabIndex = 0
        Me.EPN_NO_Label.Text = "-"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 33)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(857, 291)
        Me.DataGridView1.TabIndex = 1
        '
        'Clash_Check_Create_Lebel
        '
        Me.Clash_Check_Create_Lebel.AutoSize = True
        Me.Clash_Check_Create_Lebel.Location = New System.Drawing.Point(18, 338)
        Me.Clash_Check_Create_Lebel.Name = "Clash_Check_Create_Lebel"
        Me.Clash_Check_Create_Lebel.Size = New System.Drawing.Size(29, 12)
        Me.Clash_Check_Create_Lebel.TabIndex = 0
        Me.Clash_Check_Create_Lebel.Text = "-123"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button1.Location = New System.Drawing.Point(746, 330)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(123, 29)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "확인"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ZA_Clash_Check_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(881, 366)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Clash_Check_Create_Lebel)
        Me.Controls.Add(Me.EPN_NO_Label)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ZA_Clash_Check_Create"
        Me.Text = "Clash Check"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents EPN_NO_Label As System.Windows.Forms.Label
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Clash_Check_Create_Lebel As System.Windows.Forms.Label
    Public WithEvents Button1 As System.Windows.Forms.Button
End Class
