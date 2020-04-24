<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class C_Coupling_Cover_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(C_Coupling_Cover_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.Coupling_Cover_Clash_Check = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Coupling_Cover_Create = New System.Windows.Forms.Button()
        Me.UI_STAGE_COMBO = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.Data_Path_List = New System.Windows.Forms.ListBox()
        Me.UI_PUNCH_HOLE = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ProgressBar1)
        Me.GroupBox1.Controls.Add(Me.cancel_button)
        Me.GroupBox1.Controls.Add(Me.Coupling_Cover_Clash_Check)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Message_Label)
        Me.GroupBox1.Location = New System.Drawing.Point(0, -8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(380, 433)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 402)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(355, 17)
        Me.ProgressBar1.TabIndex = 6
        '
        'cancel_button
        '
        Me.cancel_button.Enabled = False
        Me.cancel_button.Location = New System.Drawing.Point(234, 363)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(133, 33)
        Me.cancel_button.TabIndex = 3
        Me.cancel_button.Text = "종료"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'Coupling_Cover_Clash_Check
        '
        Me.Coupling_Cover_Clash_Check.Location = New System.Drawing.Point(234, 324)
        Me.Coupling_Cover_Clash_Check.Name = "Coupling_Cover_Clash_Check"
        Me.Coupling_Cover_Clash_Check.Size = New System.Drawing.Size(133, 33)
        Me.Coupling_Cover_Clash_Check.TabIndex = 3
        Me.Coupling_Cover_Clash_Check.Text = "간섭 Check"
        Me.Coupling_Cover_Clash_Check.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 138)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(355, 180)
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
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Coupling_Cover_Create)
        Me.GroupBox2.Controls.Add(Me.UI_STAGE_COMBO)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 20)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(355, 112)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Coupling Cover 실적 형상 검색"
        '
        'Coupling_Cover_Create
        '
        Me.Coupling_Cover_Create.Location = New System.Drawing.Point(216, 69)
        Me.Coupling_Cover_Create.Name = "Coupling_Cover_Create"
        Me.Coupling_Cover_Create.Size = New System.Drawing.Size(133, 33)
        Me.Coupling_Cover_Create.TabIndex = 3
        Me.Coupling_Cover_Create.Text = "실적 형상 검색"
        Me.Coupling_Cover_Create.UseVisualStyleBackColor = True
        '
        'UI_STAGE_COMBO
        '
        Me.UI_STAGE_COMBO.FormattingEnabled = True
        Me.UI_STAGE_COMBO.Location = New System.Drawing.Point(66, 76)
        Me.UI_STAGE_COMBO.Name = "UI_STAGE_COMBO"
        Me.UI_STAGE_COMBO.Size = New System.Drawing.Size(104, 20)
        Me.UI_STAGE_COMBO.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 12)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "STAGE :"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.VVIP_Check)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 20)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(343, 43)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(21, 17)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(129, 16)
        Me.VVIP_Check.TabIndex = 0
        Me.VVIP_Check.Text = "VVIP OPTION 선택"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'Message_Label
        '
        Me.Message_Label.AutoSize = True
        Me.Message_Label.Location = New System.Drawing.Point(16, 373)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(11, 12)
        Me.Message_Label.TabIndex = 1
        Me.Message_Label.Text = "-"
        '
        'Data_Path_List
        '
        Me.Data_Path_List.FormattingEnabled = True
        Me.Data_Path_List.ItemHeight = 12
        Me.Data_Path_List.Location = New System.Drawing.Point(12, 449)
        Me.Data_Path_List.Name = "Data_Path_List"
        Me.Data_Path_List.Size = New System.Drawing.Size(355, 40)
        Me.Data_Path_List.TabIndex = 2
        '
        'UI_PUNCH_HOLE
        '
        Me.UI_PUNCH_HOLE.AutoSize = True
        Me.UI_PUNCH_HOLE.Checked = True
        Me.UI_PUNCH_HOLE.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UI_PUNCH_HOLE.Location = New System.Drawing.Point(398, 23)
        Me.UI_PUNCH_HOLE.Name = "UI_PUNCH_HOLE"
        Me.UI_PUNCH_HOLE.Size = New System.Drawing.Size(183, 16)
        Me.UI_PUNCH_HOLE.TabIndex = 3
        Me.UI_PUNCH_HOLE.Text = "OPTIN 적용 (타공 유무 포함)"
        Me.UI_PUNCH_HOLE.UseVisualStyleBackColor = True
        '
        'C_Coupling_Cover_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(591, 503)
        Me.Controls.Add(Me.UI_PUNCH_HOLE)
        Me.Controls.Add(Me.Data_Path_List)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "C_Coupling_Cover_Create"
        Me.Text = "C_Coupling_Cover_Create"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Coupling_Cover_Create As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents Coupling_Cover_Clash_Check As System.Windows.Forms.Button
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents Data_Path_List As System.Windows.Forms.ListBox
    Public WithEvents UI_STAGE_COMBO As System.Windows.Forms.ComboBox
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents UI_PUNCH_HOLE As System.Windows.Forms.CheckBox
End Class
