<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_AC_Outlet_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(F_AC_Outlet_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cancel_button = New System.Windows.Forms.Button()
        Me.AC_Outlet_Clash_Check = New System.Windows.Forms.Button()
        Me.AC_Outlet_Create = New System.Windows.Forms.Button()
        Me.Data_Path_List = New System.Windows.Forms.ListBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.Check1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ProgressBar1)
        Me.GroupBox1.Controls.Add(Me.Message_Label)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.cancel_button)
        Me.GroupBox1.Controls.Add(Me.AC_Outlet_Clash_Check)
        Me.GroupBox1.Controls.Add(Me.AC_Outlet_Create)
        Me.GroupBox1.Location = New System.Drawing.Point(0, -8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(542, 393)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(14, 363)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(360, 17)
        Me.ProgressBar1.TabIndex = 10
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(12, 339)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(362, 21)
        Me.Message_Label.TabIndex = 9
        Me.Message_Label.Text = "-"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 65)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(518, 212)
        Me.DataGridView1.TabIndex = 7
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
        'cancel_button
        '
        Me.cancel_button.Enabled = False
        Me.cancel_button.Location = New System.Drawing.Point(389, 339)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.Size = New System.Drawing.Size(141, 41)
        Me.cancel_button.TabIndex = 0
        Me.cancel_button.Text = "종료"
        Me.cancel_button.UseVisualStyleBackColor = True
        '
        'AC_Outlet_Clash_Check
        '
        Me.AC_Outlet_Clash_Check.Location = New System.Drawing.Point(389, 292)
        Me.AC_Outlet_Clash_Check.Name = "AC_Outlet_Clash_Check"
        Me.AC_Outlet_Clash_Check.Size = New System.Drawing.Size(141, 41)
        Me.AC_Outlet_Clash_Check.TabIndex = 0
        Me.AC_Outlet_Clash_Check.Text = "간섭 Check"
        Me.AC_Outlet_Clash_Check.UseVisualStyleBackColor = True
        '
        'AC_Outlet_Create
        '
        Me.AC_Outlet_Create.Location = New System.Drawing.Point(389, 18)
        Me.AC_Outlet_Create.Name = "AC_Outlet_Create"
        Me.AC_Outlet_Create.Size = New System.Drawing.Size(141, 41)
        Me.AC_Outlet_Create.TabIndex = 0
        Me.AC_Outlet_Create.Text = "실적 형상 검색"
        Me.AC_Outlet_Create.UseVisualStyleBackColor = True
        '
        'Data_Path_List
        '
        Me.Data_Path_List.FormattingEnabled = True
        Me.Data_Path_List.ItemHeight = 12
        Me.Data_Path_List.Location = New System.Drawing.Point(14, 415)
        Me.Data_Path_List.Name = "Data_Path_List"
        Me.Data_Path_List.Size = New System.Drawing.Size(516, 40)
        Me.Data_Path_List.TabIndex = 3
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.VVIP_Check)
        Me.GroupBox2.Controls.Add(Me.Check1)
        Me.GroupBox2.Location = New System.Drawing.Point(558, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(131, 373)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "VVIP OPTION"
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(21, 126)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(79, 16)
        Me.VVIP_Check.TabIndex = 0
        Me.VVIP_Check.Text = "VVIP 적용"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        Me.VVIP_Check.Visible = False
        '
        'Check1
        '
        Me.Check1.AutoSize = True
        Me.Check1.Checked = True
        Me.Check1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Check1.Location = New System.Drawing.Point(21, 29)
        Me.Check1.Name = "Check1"
        Me.Check1.Size = New System.Drawing.Size(96, 16)
        Me.Check1.TabIndex = 0
        Me.Check1.Text = "10K FLANGE"
        Me.Check1.UseVisualStyleBackColor = True
        '
        'F_AC_Outlet_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(707, 474)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Data_Path_List)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "F_AC_Outlet_Create"
        Me.Text = "F_AC_Outlet_Create"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Col1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents Col5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents Data_Path_List As System.Windows.Forms.ListBox
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents AC_Outlet_Create As System.Windows.Forms.Button
    Public WithEvents cancel_button As System.Windows.Forms.Button
    Public WithEvents AC_Outlet_Clash_Check As System.Windows.Forms.Button
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents Check1 As System.Windows.Forms.CheckBox
End Class
