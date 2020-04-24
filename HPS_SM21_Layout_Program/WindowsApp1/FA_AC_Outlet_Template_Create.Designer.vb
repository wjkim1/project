<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FA_AC_Outlet_Template_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FA_AC_Outlet_Template_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AC_Outlet_Template_Create = New System.Windows.Forms.Button()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.AC_Outlet_Parameter_Control = New System.Windows.Forms.Button()
        Me.AC_Outlet_Result_Part_Create = New System.Windows.Forms.Button()
        Me.AC_Outlet_Drawing_Create = New System.Windows.Forms.Button()
        Me.AC_Outlet_Clash_Check = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.VScroll2 = New System.Windows.Forms.VScrollBar()
        Me.VScroll1 = New System.Windows.Forms.VScrollBar()
        Me.BOV_Rotate_Text = New System.Windows.Forms.TextBox()
        Me.Count_Rotate_Text = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Type_Combo = New System.Windows.Forms.ComboBox()
        Me.Stage_Combo = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.AC_Option7 = New System.Windows.Forms.RadioButton()
        Me.AC_Option10 = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Template_Part_Number_Text = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.Coupling_Cover_Parameter_Update = New System.Windows.Forms.Button()
        Me.Auto_Clash_Check = New System.Windows.Forms.CheckBox()
        Me.VVIP_Option_Check = New System.Windows.Forms.CheckBox()
        Me.Check1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.AC_Outlet_Template_Create)
        Me.GroupBox1.Controls.Add(Me.GroupBox6)
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(0, -8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(458, 666)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 253)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(434, 195)
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
        'AC_Outlet_Template_Create
        '
        Me.AC_Outlet_Template_Create.Location = New System.Drawing.Point(300, 132)
        Me.AC_Outlet_Template_Create.Name = "AC_Outlet_Template_Create"
        Me.AC_Outlet_Template_Create.Size = New System.Drawing.Size(146, 43)
        Me.AC_Outlet_Template_Create.TabIndex = 1
        Me.AC_Outlet_Template_Create.Text = "1. Template 실행"
        Me.AC_Outlet_Template_Create.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.ProgressBar1)
        Me.GroupBox6.Controls.Add(Me.Message_Label)
        Me.GroupBox6.Controls.Add(Me.AC_Outlet_Parameter_Control)
        Me.GroupBox6.Controls.Add(Me.AC_Outlet_Result_Part_Create)
        Me.GroupBox6.Controls.Add(Me.AC_Outlet_Drawing_Create)
        Me.GroupBox6.Controls.Add(Me.AC_Outlet_Clash_Check)
        Me.GroupBox6.Location = New System.Drawing.Point(12, 454)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(434, 202)
        Me.GroupBox6.TabIndex = 0
        Me.GroupBox6.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(13, 169)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(408, 17)
        Me.ProgressBar1.TabIndex = 10
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(11, 143)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(248, 18)
        Me.Message_Label.TabIndex = 9
        Me.Message_Label.Text = "-"
        '
        'AC_Outlet_Parameter_Control
        '
        Me.AC_Outlet_Parameter_Control.Location = New System.Drawing.Point(13, 21)
        Me.AC_Outlet_Parameter_Control.Name = "AC_Outlet_Parameter_Control"
        Me.AC_Outlet_Parameter_Control.Size = New System.Drawing.Size(146, 43)
        Me.AC_Outlet_Parameter_Control.TabIndex = 2
        Me.AC_Outlet_Parameter_Control.Text = "Parameter Control"
        Me.AC_Outlet_Parameter_Control.UseVisualStyleBackColor = True
        '
        'AC_Outlet_Result_Part_Create
        '
        Me.AC_Outlet_Result_Part_Create.Location = New System.Drawing.Point(275, 118)
        Me.AC_Outlet_Result_Part_Create.Name = "AC_Outlet_Result_Part_Create"
        Me.AC_Outlet_Result_Part_Create.Size = New System.Drawing.Size(146, 43)
        Me.AC_Outlet_Result_Part_Create.TabIndex = 2
        Me.AC_Outlet_Result_Part_Create.Text = "3. PART 확정"
        Me.AC_Outlet_Result_Part_Create.UseVisualStyleBackColor = True
        '
        'AC_Outlet_Drawing_Create
        '
        Me.AC_Outlet_Drawing_Create.Location = New System.Drawing.Point(275, 69)
        Me.AC_Outlet_Drawing_Create.Name = "AC_Outlet_Drawing_Create"
        Me.AC_Outlet_Drawing_Create.Size = New System.Drawing.Size(146, 43)
        Me.AC_Outlet_Drawing_Create.TabIndex = 2
        Me.AC_Outlet_Drawing_Create.Text = "2. 도면 생성"
        Me.AC_Outlet_Drawing_Create.UseVisualStyleBackColor = True
        '
        'AC_Outlet_Clash_Check
        '
        Me.AC_Outlet_Clash_Check.Location = New System.Drawing.Point(275, 20)
        Me.AC_Outlet_Clash_Check.Name = "AC_Outlet_Clash_Check"
        Me.AC_Outlet_Clash_Check.Size = New System.Drawing.Size(146, 43)
        Me.AC_Outlet_Clash_Check.TabIndex = 2
        Me.AC_Outlet_Clash_Check.Text = "간섭 Check"
        Me.AC_Outlet_Clash_Check.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.VScroll2)
        Me.GroupBox5.Controls.Add(Me.VScroll1)
        Me.GroupBox5.Controls.Add(Me.BOV_Rotate_Text)
        Me.GroupBox5.Controls.Add(Me.Count_Rotate_Text)
        Me.GroupBox5.Controls.Add(Me.Label6)
        Me.GroupBox5.Controls.Add(Me.Label5)
        Me.GroupBox5.Location = New System.Drawing.Point(12, 181)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(434, 66)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Parameter Value Control"
        '
        'VScroll2
        '
        Me.VScroll2.Location = New System.Drawing.Point(401, 16)
        Me.VScroll2.Name = "VScroll2"
        Me.VScroll2.Size = New System.Drawing.Size(20, 38)
        Me.VScroll2.TabIndex = 3
        '
        'VScroll1
        '
        Me.VScroll1.Location = New System.Drawing.Point(209, 17)
        Me.VScroll1.Name = "VScroll1"
        Me.VScroll1.Size = New System.Drawing.Size(20, 38)
        Me.VScroll1.TabIndex = 3
        '
        'BOV_Rotate_Text
        '
        Me.BOV_Rotate_Text.Location = New System.Drawing.Point(318, 26)
        Me.BOV_Rotate_Text.Name = "BOV_Rotate_Text"
        Me.BOV_Rotate_Text.Size = New System.Drawing.Size(80, 21)
        Me.BOV_Rotate_Text.TabIndex = 2
        '
        'Count_Rotate_Text
        '
        Me.Count_Rotate_Text.Location = New System.Drawing.Point(126, 26)
        Me.Count_Rotate_Text.Name = "Count_Rotate_Text"
        Me.Count_Rotate_Text.Size = New System.Drawing.Size(80, 21)
        Me.Count_Rotate_Text.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(246, 31)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(66, 12)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "BOV 회전 :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 30)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(107, 12)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Block Valve 회전 :"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Type_Combo)
        Me.GroupBox3.Controls.Add(Me.Stage_Combo)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 69)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(434, 50)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        '
        'Type_Combo
        '
        Me.Type_Combo.FormattingEnabled = True
        Me.Type_Combo.Location = New System.Drawing.Point(286, 18)
        Me.Type_Combo.Name = "Type_Combo"
        Me.Type_Combo.Size = New System.Drawing.Size(90, 20)
        Me.Type_Combo.TabIndex = 2
        '
        'Stage_Combo
        '
        Me.Stage_Combo.FormattingEnabled = True
        Me.Stage_Combo.Location = New System.Drawing.Point(110, 18)
        Me.Stage_Combo.Name = "Stage_Combo"
        Me.Stage_Combo.Size = New System.Drawing.Size(90, 20)
        Me.Stage_Combo.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(235, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 12)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "TYPE :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(50, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "STAGE :"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.AC_Option7)
        Me.GroupBox4.Controls.Add(Me.AC_Option10)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Location = New System.Drawing.Point(12, 125)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(280, 50)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        '
        'AC_Option7
        '
        Me.AC_Option7.AutoSize = True
        Me.AC_Option7.Location = New System.Drawing.Point(190, 20)
        Me.AC_Option7.Name = "AC_Option7"
        Me.AC_Option7.Size = New System.Drawing.Size(41, 16)
        Me.AC_Option7.TabIndex = 1
        Me.AC_Option7.Text = "NO"
        Me.AC_Option7.UseVisualStyleBackColor = True
        '
        'AC_Option10
        '
        Me.AC_Option10.AutoSize = True
        Me.AC_Option10.Checked = True
        Me.AC_Option10.Location = New System.Drawing.Point(123, 20)
        Me.AC_Option10.Name = "AC_Option10"
        Me.AC_Option10.Size = New System.Drawing.Size(47, 16)
        Me.AC_Option10.TabIndex = 1
        Me.AC_Option10.TabStop = True
        Me.AC_Option10.Text = "YES"
        Me.AC_Option10.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 12)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Silencer 유/무 :"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Template_Part_Number_Text)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 13)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(434, 50)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'Template_Part_Number_Text
        '
        Me.Template_Part_Number_Text.Location = New System.Drawing.Point(172, 17)
        Me.Template_Part_Number_Text.Name = "Template_Part_Number_Text"
        Me.Template_Part_Number_Text.Size = New System.Drawing.Size(244, 21)
        Me.Template_Part_Number_Text.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(155, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "AC Outlet Pipe PART No. :"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.Coupling_Cover_Parameter_Update)
        Me.GroupBox7.Controls.Add(Me.Auto_Clash_Check)
        Me.GroupBox7.Controls.Add(Me.VVIP_Option_Check)
        Me.GroupBox7.Controls.Add(Me.Check1)
        Me.GroupBox7.Location = New System.Drawing.Point(464, 12)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(172, 646)
        Me.GroupBox7.TabIndex = 1
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "VVIP OPTION"
        '
        'Coupling_Cover_Parameter_Update
        '
        Me.Coupling_Cover_Parameter_Update.Location = New System.Drawing.Point(41, 59)
        Me.Coupling_Cover_Parameter_Update.Name = "Coupling_Cover_Parameter_Update"
        Me.Coupling_Cover_Parameter_Update.Size = New System.Drawing.Size(116, 36)
        Me.Coupling_Cover_Parameter_Update.TabIndex = 2
        Me.Coupling_Cover_Parameter_Update.Text = "Update"
        Me.Coupling_Cover_Parameter_Update.UseVisualStyleBackColor = True
        '
        'Auto_Clash_Check
        '
        Me.Auto_Clash_Check.AutoSize = True
        Me.Auto_Clash_Check.Checked = True
        Me.Auto_Clash_Check.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Auto_Clash_Check.Location = New System.Drawing.Point(22, 186)
        Me.Auto_Clash_Check.Name = "Auto_Clash_Check"
        Me.Auto_Clash_Check.Size = New System.Drawing.Size(128, 16)
        Me.Auto_Clash_Check.TabIndex = 0
        Me.Auto_Clash_Check.Text = "간섭체크 자동 확인"
        Me.Auto_Clash_Check.UseVisualStyleBackColor = True
        Me.Auto_Clash_Check.Visible = False
        '
        'VVIP_Option_Check
        '
        Me.VVIP_Option_Check.AutoSize = True
        Me.VVIP_Option_Check.Location = New System.Drawing.Point(22, 125)
        Me.VVIP_Option_Check.Name = "VVIP_Option_Check"
        Me.VVIP_Option_Check.Size = New System.Drawing.Size(79, 16)
        Me.VVIP_Option_Check.TabIndex = 0
        Me.VVIP_Option_Check.Text = "VVIP 적용"
        Me.VVIP_Option_Check.UseVisualStyleBackColor = True
        Me.VVIP_Option_Check.Visible = False
        '
        'Check1
        '
        Me.Check1.AutoSize = True
        Me.Check1.Location = New System.Drawing.Point(22, 31)
        Me.Check1.Name = "Check1"
        Me.Check1.Size = New System.Drawing.Size(96, 16)
        Me.Check1.TabIndex = 0
        Me.Check1.Text = "10K FLANGE"
        Me.Check1.UseVisualStyleBackColor = True
        '
        'FA_AC_Outlet_Template_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(678, 657)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FA_AC_Outlet_Template_Create"
        Me.Text = "FA_AC_Outlet_Template_Create"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
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
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents AC_Outlet_Template_Create As System.Windows.Forms.Button
    Public WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Public WithEvents VScroll2 As System.Windows.Forms.VScrollBar
    Public WithEvents VScroll1 As System.Windows.Forms.VScrollBar
    Public WithEvents BOV_Rotate_Text As System.Windows.Forms.TextBox
    Public WithEvents Count_Rotate_Text As System.Windows.Forms.TextBox
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents Type_Combo As System.Windows.Forms.ComboBox
    Public WithEvents Stage_Combo As System.Windows.Forms.ComboBox
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents AC_Option7 As System.Windows.Forms.RadioButton
    Public WithEvents AC_Option10 As System.Windows.Forms.RadioButton
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents Template_Part_Number_Text As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents AC_Outlet_Parameter_Control As System.Windows.Forms.Button
    Public WithEvents AC_Outlet_Result_Part_Create As System.Windows.Forms.Button
    Public WithEvents AC_Outlet_Drawing_Create As System.Windows.Forms.Button
    Public WithEvents AC_Outlet_Clash_Check As System.Windows.Forms.Button
    Public WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Public WithEvents Coupling_Cover_Parameter_Update As System.Windows.Forms.Button
    Public WithEvents VVIP_Option_Check As System.Windows.Forms.CheckBox
    Public WithEvents Check1 As System.Windows.Forms.CheckBox
    Public WithEvents Auto_Clash_Check As System.Windows.Forms.CheckBox
End Class
