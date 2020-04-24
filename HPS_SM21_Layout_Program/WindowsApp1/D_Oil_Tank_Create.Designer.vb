<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D_Oil_Tank_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D_Oil_Tank_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.Oil_Tank_Clash_Check = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Oil_Tank_Create = New System.Windows.Forms.Button()
        Me.Option_CHN = New System.Windows.Forms.RadioButton()
        Me.Option_KOR = New System.Windows.Forms.RadioButton()
        Me.Option_Default = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Data_Path_List = New System.Windows.Forms.ListBox()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.UI_PPN05_D12 = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ProgressBar1)
        Me.GroupBox1.Controls.Add(Me.Message_Label)
        Me.GroupBox1.Controls.Add(Me.cancel_button)
        Me.GroupBox1.Controls.Add(Me.Oil_Tank_Clash_Check)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(0, -8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(696, 450)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 422)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(672, 14)
        Me.ProgressBar1.TabIndex = 8
        '
        'Message_Label
        '
        Me.Message_Label.AutoSize = True
        Me.Message_Label.Location = New System.Drawing.Point(28, 392)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(11, 12)
        Me.Message_Label.TabIndex = 7
        Me.Message_Label.Text = "-"
        '
        'cancel_button
        '
        Me.cancel_button.Enabled = False
        Me.cancel_button.Location = New System.Drawing.Point(559, 380)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(125, 36)
        Me.cancel_button.TabIndex = 6
        Me.cancel_button.Text = "종료"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'Oil_Tank_Clash_Check
        '
        Me.Oil_Tank_Clash_Check.Location = New System.Drawing.Point(559, 338)
        Me.Oil_Tank_Clash_Check.Name = "Oil_Tank_Clash_Check"
        Me.Oil_Tank_Clash_Check.Size = New System.Drawing.Size(125, 36)
        Me.Oil_Tank_Clash_Check.TabIndex = 6
        Me.Oil_Tank_Clash_Check.Text = "간섭 Check"
        Me.Oil_Tank_Clash_Check.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 96)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(672, 236)
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
        Me.GroupBox2.Controls.Add(Me.Oil_Tank_Create)
        Me.GroupBox2.Controls.Add(Me.Option_CHN)
        Me.GroupBox2.Controls.Add(Me.Option_KOR)
        Me.GroupBox2.Controls.Add(Me.Option_Default)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 20)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(672, 70)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Oil Tank 실적 형상 검색"
        '
        'Oil_Tank_Create
        '
        Me.Oil_Tank_Create.Location = New System.Drawing.Point(541, 20)
        Me.Oil_Tank_Create.Name = "Oil_Tank_Create"
        Me.Oil_Tank_Create.Size = New System.Drawing.Size(125, 36)
        Me.Oil_Tank_Create.TabIndex = 2
        Me.Oil_Tank_Create.Text = "실적 형상 검색"
        Me.Oil_Tank_Create.UseVisualStyleBackColor = True
        '
        'Option_CHN
        '
        Me.Option_CHN.AutoSize = True
        Me.Option_CHN.Location = New System.Drawing.Point(284, 30)
        Me.Option_CHN.Name = "Option_CHN"
        Me.Option_CHN.Size = New System.Drawing.Size(49, 16)
        Me.Option_CHN.TabIndex = 1
        Me.Option_CHN.Text = "CHN"
        Me.Option_CHN.UseVisualStyleBackColor = True
        '
        'Option_KOR
        '
        Me.Option_KOR.AutoSize = True
        Me.Option_KOR.Location = New System.Drawing.Point(214, 30)
        Me.Option_KOR.Name = "Option_KOR"
        Me.Option_KOR.Size = New System.Drawing.Size(48, 16)
        Me.Option_KOR.TabIndex = 1
        Me.Option_KOR.Text = "KOR"
        Me.Option_KOR.UseVisualStyleBackColor = True
        '
        'Option_Default
        '
        Me.Option_Default.AutoSize = True
        Me.Option_Default.Checked = True
        Me.Option_Default.Location = New System.Drawing.Point(132, 30)
        Me.Option_Default.Name = "Option_Default"
        Me.Option_Default.Size = New System.Drawing.Size(61, 16)
        Me.Option_Default.TabIndex = 1
        Me.Option_Default.TabStop = True
        Me.Option_Default.Text = "Default"
        Me.Option_Default.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Oil Tank Type :"
        '
        'Data_Path_List
        '
        Me.Data_Path_List.FormattingEnabled = True
        Me.Data_Path_List.ItemHeight = 12
        Me.Data_Path_List.Location = New System.Drawing.Point(12, 457)
        Me.Data_Path_List.Name = "Data_Path_List"
        Me.Data_Path_List.Size = New System.Drawing.Size(472, 40)
        Me.Data_Path_List.TabIndex = 1
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Checked = True
        Me.VVIP_Check.CheckState = System.Windows.Forms.CheckState.Checked
        Me.VVIP_Check.Location = New System.Drawing.Point(715, 62)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(129, 16)
        Me.VVIP_Check.TabIndex = 2
        Me.VVIP_Check.Text = "VVIP OPTION 선택"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'UI_PPN05_D12
        '
        Me.UI_PPN05_D12.AutoSize = True
        Me.UI_PPN05_D12.Checked = True
        Me.UI_PPN05_D12.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UI_PPN05_D12.Location = New System.Drawing.Point(715, 23)
        Me.UI_PPN05_D12.Name = "UI_PPN05_D12"
        Me.UI_PPN05_D12.Size = New System.Drawing.Size(139, 16)
        Me.UI_PPN05_D12.TabIndex = 3
        Me.UI_PPN05_D12.Text = "TRANSMITTER 적용"
        Me.UI_PPN05_D12.UseVisualStyleBackColor = True
        '
        'D_Oil_Tank_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(868, 504)
        Me.Controls.Add(Me.VVIP_Check)
        Me.Controls.Add(Me.UI_PPN05_D12)
        Me.Controls.Add(Me.Data_Path_List)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "D_Oil_Tank_Create"
        Me.Text = "D_Oil_Tank_Create"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents Option_CHN As System.Windows.Forms.RadioButton
    Public WithEvents Option_KOR As System.Windows.Forms.RadioButton
    Public WithEvents Option_Default As System.Windows.Forms.RadioButton
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Oil_Tank_Create As System.Windows.Forms.Button
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents cancel_button As System.Windows.Forms.Button
    Public WithEvents Oil_Tank_Clash_Check As System.Windows.Forms.Button
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Data_Path_List As System.Windows.Forms.ListBox
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents UI_PPN05_D12 As System.Windows.Forms.CheckBox
End Class
