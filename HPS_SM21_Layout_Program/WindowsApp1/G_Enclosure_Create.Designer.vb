<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class G_Enclosure_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(G_Enclosure_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.Enclosure_Clash_Check = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Enclosure_Create = New System.Windows.Forms.Button()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.VVIP_Frame = New System.Windows.Forms.GroupBox()
        Me.UI_EPN11_D14 = New System.Windows.Forms.CheckBox()
        Me.UI_EPN11_D13 = New System.Windows.Forms.CheckBox()
        Me.UI_EPN11_D12 = New System.Windows.Forms.CheckBox()
        Me.Data_Path_List = New System.Windows.Forms.ListBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.VVIP_Frame.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ProgressBar1)
        Me.GroupBox1.Controls.Add(Me.Message_Label)
        Me.GroupBox1.Controls.Add(Me.cancel_button)
        Me.GroupBox1.Controls.Add(Me.Enclosure_Clash_Check)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Location = New System.Drawing.Point(0, -8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(460, 421)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 388)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(432, 17)
        Me.ProgressBar1.TabIndex = 12
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(10, 359)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(269, 12)
        Me.Message_Label.TabIndex = 11
        Me.Message_Label.Text = "- "
        '
        'cancel_button
        '
        Me.cancel_button.Enabled = False
        Me.cancel_button.Location = New System.Drawing.Point(285, 348)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(159, 34)
        Me.cancel_button.TabIndex = 10
        Me.cancel_button.Text = "종료"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'Enclosure_Clash_Check
        '
        Me.Enclosure_Clash_Check.Location = New System.Drawing.Point(285, 308)
        Me.Enclosure_Clash_Check.Name = "Enclosure_Clash_Check"
        Me.Enclosure_Clash_Check.Size = New System.Drawing.Size(159, 34)
        Me.Enclosure_Clash_Check.TabIndex = 10
        Me.Enclosure_Clash_Check.Text = "간섭 Check"
        Me.Enclosure_Clash_Check.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 95)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(432, 195)
        Me.DataGridView1.TabIndex = 9
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
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Enclosure_Create)
        Me.GroupBox3.Controls.Add(Me.VVIP_Check)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 20)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(432, 69)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Enclosure 실적 형상 검색"
        '
        'Enclosure_Create
        '
        Me.Enclosure_Create.Location = New System.Drawing.Point(254, 21)
        Me.Enclosure_Create.Name = "Enclosure_Create"
        Me.Enclosure_Create.Size = New System.Drawing.Size(159, 34)
        Me.Enclosure_Create.TabIndex = 1
        Me.Enclosure_Create.Text = "실적 형상 검색"
        Me.Enclosure_Create.UseVisualStyleBackColor = True
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(17, 31)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(129, 16)
        Me.VVIP_Check.TabIndex = 0
        Me.VVIP_Check.Text = "VVIP OPTION 선택"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'VVIP_Frame
        '
        Me.VVIP_Frame.Controls.Add(Me.UI_EPN11_D14)
        Me.VVIP_Frame.Controls.Add(Me.UI_EPN11_D13)
        Me.VVIP_Frame.Controls.Add(Me.UI_EPN11_D12)
        Me.VVIP_Frame.Location = New System.Drawing.Point(486, -8)
        Me.VVIP_Frame.Name = "VVIP_Frame"
        Me.VVIP_Frame.Size = New System.Drawing.Size(192, 421)
        Me.VVIP_Frame.TabIndex = 0
        Me.VVIP_Frame.TabStop = False
        '
        'UI_EPN11_D14
        '
        Me.UI_EPN11_D14.AutoSize = True
        Me.UI_EPN11_D14.Checked = True
        Me.UI_EPN11_D14.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UI_EPN11_D14.Location = New System.Drawing.Point(22, 116)
        Me.UI_EPN11_D14.Name = "UI_EPN11_D14"
        Me.UI_EPN11_D14.Size = New System.Drawing.Size(155, 16)
        Me.UI_EPN11_D14.TabIndex = 1
        Me.UI_EPN11_D14.Text = "CONTROL PANEL 외장"
        Me.UI_EPN11_D14.UseVisualStyleBackColor = True
        '
        'UI_EPN11_D13
        '
        Me.UI_EPN11_D13.AutoSize = True
        Me.UI_EPN11_D13.Checked = True
        Me.UI_EPN11_D13.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UI_EPN11_D13.Location = New System.Drawing.Point(22, 85)
        Me.UI_EPN11_D13.Name = "UI_EPN11_D13"
        Me.UI_EPN11_D13.Size = New System.Drawing.Size(154, 16)
        Me.UI_EPN11_D13.TabIndex = 1
        Me.UI_EPN11_D13.Text = "VVIP DOOR LOCK 적용"
        Me.UI_EPN11_D13.UseVisualStyleBackColor = True
        '
        'UI_EPN11_D12
        '
        Me.UI_EPN11_D12.AutoSize = True
        Me.UI_EPN11_D12.Checked = True
        Me.UI_EPN11_D12.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UI_EPN11_D12.Location = New System.Drawing.Point(22, 51)
        Me.UI_EPN11_D12.Name = "UI_EPN11_D12"
        Me.UI_EPN11_D12.Size = New System.Drawing.Size(103, 16)
        Me.UI_EPN11_D12.TabIndex = 1
        Me.UI_EPN11_D12.Text = "CANOPY 적용"
        Me.UI_EPN11_D12.UseVisualStyleBackColor = True
        '
        'Data_Path_List
        '
        Me.Data_Path_List.FormattingEnabled = True
        Me.Data_Path_List.ItemHeight = 12
        Me.Data_Path_List.Location = New System.Drawing.Point(0, 444)
        Me.Data_Path_List.Name = "Data_Path_List"
        Me.Data_Path_List.Size = New System.Drawing.Size(460, 40)
        Me.Data_Path_List.TabIndex = 4
        '
        'G_Enclosure_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(702, 514)
        Me.Controls.Add(Me.Data_Path_List)
        Me.Controls.Add(Me.VVIP_Frame)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "G_Enclosure_Create"
        Me.Text = "G_Enclosure_Create"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.VVIP_Frame.ResumeLayout(False)
        Me.VVIP_Frame.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Data_Path_List As System.Windows.Forms.ListBox
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents Enclosure_Create As System.Windows.Forms.Button
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents VVIP_Frame As System.Windows.Forms.GroupBox
    Public WithEvents cancel_button As System.Windows.Forms.Button
    Public WithEvents Enclosure_Clash_Check As System.Windows.Forms.Button
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents UI_EPN11_D14 As System.Windows.Forms.CheckBox
    Public WithEvents UI_EPN11_D13 As System.Windows.Forms.CheckBox
    Public WithEvents UI_EPN11_D12 As System.Windows.Forms.CheckBox
End Class
