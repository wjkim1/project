<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class E_Oil_Piping_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(E_Oil_Piping_Create))
        Me.Data_Path_List = New System.Windows.Forms.ListBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.Oil_Piping_Clash_Check = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GBX_RangeGroupBox = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Range = New System.Windows.Forms.TextBox()
        Me.Range_Check = New System.Windows.Forms.CheckBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.Oil_Piping_Create = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.UI_SIGHT_GLASS = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GBX_RangeGroupBox.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Data_Path_List
        '
        Me.Data_Path_List.FormattingEnabled = True
        Me.Data_Path_List.ItemHeight = 12
        Me.Data_Path_List.Location = New System.Drawing.Point(12, 442)
        Me.Data_Path_List.Name = "Data_Path_List"
        Me.Data_Path_List.Size = New System.Drawing.Size(563, 40)
        Me.Data_Path_List.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(0, -8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(586, 420)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
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
        Me.DataGridView1.Size = New System.Drawing.Size(563, 195)
        Me.DataGridView1.TabIndex = 8
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
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.ProgressBar1)
        Me.GroupBox4.Controls.Add(Me.Message_Label)
        Me.GroupBox4.Controls.Add(Me.cancel_button)
        Me.GroupBox4.Controls.Add(Me.Oil_Piping_Clash_Check)
        Me.GroupBox4.Location = New System.Drawing.Point(12, 297)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(563, 113)
        Me.GroupBox4.TabIndex = 5
        Me.GroupBox4.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 85)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(370, 17)
        Me.ProgressBar1.TabIndex = 7
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(10, 64)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(372, 12)
        Me.Message_Label.TabIndex = 6
        Me.Message_Label.Text = "- "
        '
        'cancel_button
        '
        Me.cancel_button.Enabled = False
        Me.cancel_button.Location = New System.Drawing.Point(393, 64)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(158, 38)
        Me.cancel_button.TabIndex = 5
        Me.cancel_button.Text = "종료"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'Oil_Piping_Clash_Check
        '
        Me.Oil_Piping_Clash_Check.Location = New System.Drawing.Point(393, 20)
        Me.Oil_Piping_Clash_Check.Name = "Oil_Piping_Clash_Check"
        Me.Oil_Piping_Clash_Check.Size = New System.Drawing.Size(158, 38)
        Me.Oil_Piping_Clash_Check.TabIndex = 4
        Me.Oil_Piping_Clash_Check.Text = "간섭 Check"
        Me.Oil_Piping_Clash_Check.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.GBX_RangeGroupBox)
        Me.GroupBox2.Controls.Add(Me.GroupBox5)
        Me.GroupBox2.Controls.Add(Me.Oil_Piping_Create)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 20)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(563, 70)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Oil Piping 실적 형상 검색"
        '
        'GBX_RangeGroupBox
        '
        Me.GBX_RangeGroupBox.Controls.Add(Me.Label1)
        Me.GBX_RangeGroupBox.Controls.Add(Me.Range)
        Me.GBX_RangeGroupBox.Controls.Add(Me.Range_Check)
        Me.GBX_RangeGroupBox.Location = New System.Drawing.Point(170, 15)
        Me.GBX_RangeGroupBox.Name = "GBX_RangeGroupBox"
        Me.GBX_RangeGroupBox.Size = New System.Drawing.Size(217, 46)
        Me.GBX_RangeGroupBox.TabIndex = 9
        Me.GBX_RangeGroupBox.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(172, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "mm"
        '
        'Range
        '
        Me.Range.Location = New System.Drawing.Point(84, 18)
        Me.Range.Name = "Range"
        Me.Range.Size = New System.Drawing.Size(82, 21)
        Me.Range.TabIndex = 4
        '
        'Range_Check
        '
        Me.Range_Check.AutoSize = True
        Me.Range_Check.Location = New System.Drawing.Point(15, 20)
        Me.Range_Check.Name = "Range_Check"
        Me.Range_Check.Size = New System.Drawing.Size(68, 16)
        Me.Range_Check.TabIndex = 3
        Me.Range_Check.Text = "Range :"
        Me.Range_Check.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.VVIP_Check)
        Me.GroupBox5.Location = New System.Drawing.Point(12, 15)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(152, 46)
        Me.GroupBox5.TabIndex = 7
        Me.GroupBox5.TabStop = False
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(15, 20)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(129, 16)
        Me.VVIP_Check.TabIndex = 1
        Me.VVIP_Check.Text = "VVIP OPTION 선택"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'Oil_Piping_Create
        '
        Me.Oil_Piping_Create.Location = New System.Drawing.Point(393, 22)
        Me.Oil_Piping_Create.Name = "Oil_Piping_Create"
        Me.Oil_Piping_Create.Size = New System.Drawing.Size(158, 38)
        Me.Oil_Piping_Create.TabIndex = 3
        Me.Oil_Piping_Create.Text = "실적 형상 검색"
        Me.Oil_Piping_Create.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.UI_SIGHT_GLASS)
        Me.GroupBox3.Location = New System.Drawing.Point(602, 12)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(162, 400)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "VVIP OPTION"
        '
        'UI_SIGHT_GLASS
        '
        Me.UI_SIGHT_GLASS.AutoSize = True
        Me.UI_SIGHT_GLASS.Checked = True
        Me.UI_SIGHT_GLASS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UI_SIGHT_GLASS.Location = New System.Drawing.Point(18, 32)
        Me.UI_SIGHT_GLASS.Name = "UI_SIGHT_GLASS"
        Me.UI_SIGHT_GLASS.Size = New System.Drawing.Size(132, 16)
        Me.UI_SIGHT_GLASS.TabIndex = 2
        Me.UI_SIGHT_GLASS.Text = "SIGHT GLASS 적용"
        Me.UI_SIGHT_GLASS.UseVisualStyleBackColor = True
        '
        'E_Oil_Piping_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(774, 493)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Data_Path_List)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "E_Oil_Piping_Create"
        Me.Text = "E_Oil_Piping_Create"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GBX_RangeGroupBox.ResumeLayout(False)
        Me.GBX_RangeGroupBox.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents Data_Path_List As System.Windows.Forms.ListBox
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents GBX_RangeGroupBox As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Range As System.Windows.Forms.TextBox
    Public WithEvents Range_Check As System.Windows.Forms.CheckBox
    Public WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents Oil_Piping_Create As System.Windows.Forms.Button
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents cancel_button As System.Windows.Forms.Button
    Public WithEvents Oil_Piping_Clash_Check As System.Windows.Forms.Button
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents UI_SIGHT_GLASS As System.Windows.Forms.CheckBox
End Class
