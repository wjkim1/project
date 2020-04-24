<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ZB_Template_Parameter_Control
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ZB_Template_Parameter_Control))
        Me.Template_Parameter_Control_EPN_Label = New System.Windows.Forms.Label()
        Me.Template_Parameter_Control_Lebel = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Template_Parameter_Control_EPN_Label
        '
        Me.Template_Parameter_Control_EPN_Label.AutoSize = True
        Me.Template_Parameter_Control_EPN_Label.Location = New System.Drawing.Point(12, 24)
        Me.Template_Parameter_Control_EPN_Label.Name = "Template_Parameter_Control_EPN_Label"
        Me.Template_Parameter_Control_EPN_Label.Size = New System.Drawing.Size(42, 12)
        Me.Template_Parameter_Control_EPN_Label.TabIndex = 0
        Me.Template_Parameter_Control_EPN_Label.Text = "Label1"
        '
        'Template_Parameter_Control_Lebel
        '
        Me.Template_Parameter_Control_Lebel.AutoSize = True
        Me.Template_Parameter_Control_Lebel.Location = New System.Drawing.Point(12, 66)
        Me.Template_Parameter_Control_Lebel.Name = "Template_Parameter_Control_Lebel"
        Me.Template_Parameter_Control_Lebel.Size = New System.Drawing.Size(42, 12)
        Me.Template_Parameter_Control_Lebel.TabIndex = 0
        Me.Template_Parameter_Control_Lebel.Text = "Label1"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(151, 66)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(381, 167)
        Me.DataGridView1.TabIndex = 5
        '
        'Col1
        '
        Me.Col1.HeaderText = ""
        Me.Col1.Name = "Col1"
        Me.Col1.ReadOnly = True
        '
        'Col2
        '
        Me.Col2.HeaderText = ""
        Me.Col2.Name = "Col2"
        Me.Col2.ReadOnly = True
        '
        'Col3
        '
        Me.Col3.HeaderText = ""
        Me.Col3.Name = "Col3"
        Me.Col3.ReadOnly = True
        '
        'Col4
        '
        Me.Col4.HeaderText = ""
        Me.Col4.Name = "Col4"
        Me.Col4.ReadOnly = True
        '
        'Col5
        '
        Me.Col5.HeaderText = ""
        Me.Col5.Name = "Col5"
        Me.Col5.ReadOnly = True
        '
        'ZB_Template_Parameter_Control
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(888, 458)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Template_Parameter_Control_Lebel)
        Me.Controls.Add(Me.Template_Parameter_Control_EPN_Label)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ZB_Template_Parameter_Control"
        Me.Text = "ZB_Template_Parameter_Control"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Template_Parameter_Control_EPN_Label As System.Windows.Forms.Label
    Public WithEvents Template_Parameter_Control_Lebel As System.Windows.Forms.Label
End Class
