<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HA_Base_Frame_Template_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(HA_Base_Frame_Template_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Template_Part_Number_Text = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.UI_ENCLOSURE_TYPE_2 = New System.Windows.Forms.RadioButton()
        Me.UI_ENCLOSURE_TYPE = New System.Windows.Forms.RadioButton()
        Me.Option_Default = New System.Windows.Forms.RadioButton()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.Option_Over_8 = New System.Windows.Forms.RadioButton()
        Me.Option_Under_8 = New System.Windows.Forms.RadioButton()
        Me.Option_Motor_Default = New System.Windows.Forms.RadioButton()
        Me.Frm_BaseFrame_1 = New System.Windows.Forms.GroupBox()
        Me.Base_Frame_Template_Create = New System.Windows.Forms.Button()
        Me.VVIP_Check = New System.Windows.Forms.CheckBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.LEN_BASE_L_Expand_Update = New System.Windows.Forms.Button()
        Me.LEN_BASE_L_Expand_Text = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.NOW_BASE_L = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Auto_Trap_Option2 = New System.Windows.Forms.RadioButton()
        Me.Auto_Trap_Option1 = New System.Windows.Forms.RadioButton()
        Me.Frm_BaseFrame_2 = New System.Windows.Forms.GroupBox()
        Me.Base_Frame_Template_Create2 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.VVIP_Frame = New System.Windows.Forms.GroupBox()
        Me.VVIP_Message = New System.Windows.Forms.Label()
        Me.Base_Frame_Parameter_Update = New System.Windows.Forms.Button()
        Me.Auto_Clash_Check = New System.Windows.Forms.CheckBox()
        Me.UI_CONTROL_PNL = New System.Windows.Forms.CheckBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.Base_Frame_Parameter_Control = New System.Windows.Forms.Button()
        Me.Base_Frame_Result_Part_Create = New System.Windows.Forms.Button()
        Me.Base_Frame_Drawing_Create = New System.Windows.Forms.Button()
        Me.Base_Frame_Clash_Check = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frm_BaseFrame_1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.Frm_BaseFrame_2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.VVIP_Frame.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Template_Part_Number_Text)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(362, 56)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Template_Part_Number_Text
        '
        Me.Template_Part_Number_Text.Location = New System.Drawing.Point(165, 23)
        Me.Template_Part_Number_Text.Name = "Template_Part_Number_Text"
        Me.Template_Part_Number_Text.Size = New System.Drawing.Size(184, 21)
        Me.Template_Part_Number_Text.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(144, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Base Frame PART NO. :"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.UI_ENCLOSURE_TYPE_2)
        Me.GroupBox3.Controls.Add(Me.UI_ENCLOSURE_TYPE)
        Me.GroupBox3.Controls.Add(Me.Option_Default)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 136)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(362, 56)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Enclosure 유무"
        '
        'UI_ENCLOSURE_TYPE_2
        '
        Me.UI_ENCLOSURE_TYPE_2.AutoSize = True
        Me.UI_ENCLOSURE_TYPE_2.Location = New System.Drawing.Point(257, 25)
        Me.UI_ENCLOSURE_TYPE_2.Name = "UI_ENCLOSURE_TYPE_2"
        Me.UI_ENCLOSURE_TYPE_2.Size = New System.Drawing.Size(39, 16)
        Me.UI_ENCLOSURE_TYPE_2.TabIndex = 0
        Me.UI_ENCLOSURE_TYPE_2.Text = "No"
        Me.UI_ENCLOSURE_TYPE_2.UseVisualStyleBackColor = True
        '
        'UI_ENCLOSURE_TYPE
        '
        Me.UI_ENCLOSURE_TYPE.AutoSize = True
        Me.UI_ENCLOSURE_TYPE.Location = New System.Drawing.Point(127, 25)
        Me.UI_ENCLOSURE_TYPE.Name = "UI_ENCLOSURE_TYPE"
        Me.UI_ENCLOSURE_TYPE.Size = New System.Drawing.Size(45, 16)
        Me.UI_ENCLOSURE_TYPE.TabIndex = 0
        Me.UI_ENCLOSURE_TYPE.Text = "Yes"
        Me.UI_ENCLOSURE_TYPE.UseVisualStyleBackColor = True
        '
        'Option_Default
        '
        Me.Option_Default.AutoSize = True
        Me.Option_Default.Checked = True
        Me.Option_Default.Location = New System.Drawing.Point(19, 25)
        Me.Option_Default.Name = "Option_Default"
        Me.Option_Default.Size = New System.Drawing.Size(61, 16)
        Me.Option_Default.TabIndex = 0
        Me.Option_Default.TabStop = True
        Me.Option_Default.Text = "Default"
        Me.Option_Default.UseVisualStyleBackColor = True
        '
        'Frame5
        '
        Me.Frame5.Controls.Add(Me.Option_Over_8)
        Me.Frame5.Controls.Add(Me.Option_Under_8)
        Me.Frame5.Controls.Add(Me.Option_Motor_Default)
        Me.Frame5.Location = New System.Drawing.Point(12, 74)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Size = New System.Drawing.Size(362, 56)
        Me.Frame5.TabIndex = 3
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Motor 무게"
        '
        'Option_Over_8
        '
        Me.Option_Over_8.AutoSize = True
        Me.Option_Over_8.Location = New System.Drawing.Point(257, 25)
        Me.Option_Over_8.Name = "Option_Over_8"
        Me.Option_Over_8.Size = New System.Drawing.Size(92, 16)
        Me.Option_Over_8.TabIndex = 0
        Me.Option_Over_8.Text = "8,000kg 이상"
        Me.Option_Over_8.UseVisualStyleBackColor = True
        '
        'Option_Under_8
        '
        Me.Option_Under_8.AutoSize = True
        Me.Option_Under_8.Location = New System.Drawing.Point(127, 25)
        Me.Option_Under_8.Name = "Option_Under_8"
        Me.Option_Under_8.Size = New System.Drawing.Size(92, 16)
        Me.Option_Under_8.TabIndex = 0
        Me.Option_Under_8.Text = "8,000kg 미만"
        Me.Option_Under_8.UseVisualStyleBackColor = True
        '
        'Option_Motor_Default
        '
        Me.Option_Motor_Default.AutoSize = True
        Me.Option_Motor_Default.Checked = True
        Me.Option_Motor_Default.Location = New System.Drawing.Point(19, 25)
        Me.Option_Motor_Default.Name = "Option_Motor_Default"
        Me.Option_Motor_Default.Size = New System.Drawing.Size(61, 16)
        Me.Option_Motor_Default.TabIndex = 0
        Me.Option_Motor_Default.TabStop = True
        Me.Option_Motor_Default.Text = "Default"
        Me.Option_Motor_Default.UseVisualStyleBackColor = True
        '
        'Frm_BaseFrame_1
        '
        Me.Frm_BaseFrame_1.Controls.Add(Me.Base_Frame_Template_Create)
        Me.Frm_BaseFrame_1.Controls.Add(Me.VVIP_Check)
        Me.Frm_BaseFrame_1.Controls.Add(Me.GroupBox5)
        Me.Frm_BaseFrame_1.Controls.Add(Me.GroupBox4)
        Me.Frm_BaseFrame_1.Location = New System.Drawing.Point(12, 198)
        Me.Frm_BaseFrame_1.Name = "Frm_BaseFrame_1"
        Me.Frm_BaseFrame_1.Size = New System.Drawing.Size(362, 208)
        Me.Frm_BaseFrame_1.TabIndex = 4
        Me.Frm_BaseFrame_1.TabStop = False
        '
        'Base_Frame_Template_Create
        '
        Me.Base_Frame_Template_Create.Location = New System.Drawing.Point(219, 87)
        Me.Base_Frame_Template_Create.Name = "Base_Frame_Template_Create"
        Me.Base_Frame_Template_Create.Size = New System.Drawing.Size(130, 33)
        Me.Base_Frame_Template_Create.TabIndex = 4
        Me.Base_Frame_Template_Create.Text = "1. Template 실행"
        Me.Base_Frame_Template_Create.UseVisualStyleBackColor = True
        '
        'VVIP_Check
        '
        Me.VVIP_Check.AutoSize = True
        Me.VVIP_Check.Location = New System.Drawing.Point(19, 96)
        Me.VVIP_Check.Name = "VVIP_Check"
        Me.VVIP_Check.Size = New System.Drawing.Size(129, 16)
        Me.VVIP_Check.TabIndex = 3
        Me.VVIP_Check.Text = "VVIP OPTION 선택"
        Me.VVIP_Check.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.LEN_BASE_L_Expand_Update)
        Me.GroupBox5.Controls.Add(Me.LEN_BASE_L_Expand_Text)
        Me.GroupBox5.Controls.Add(Me.Label4)
        Me.GroupBox5.Controls.Add(Me.NOW_BASE_L)
        Me.GroupBox5.Controls.Add(Me.Label5)
        Me.GroupBox5.Controls.Add(Me.Label2)
        Me.GroupBox5.Location = New System.Drawing.Point(6, 126)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(350, 75)
        Me.GroupBox5.TabIndex = 2
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Base Frame 길이 조정"
        '
        'LEN_BASE_L_Expand_Update
        '
        Me.LEN_BASE_L_Expand_Update.Location = New System.Drawing.Point(282, 25)
        Me.LEN_BASE_L_Expand_Update.Name = "LEN_BASE_L_Expand_Update"
        Me.LEN_BASE_L_Expand_Update.Size = New System.Drawing.Size(61, 41)
        Me.LEN_BASE_L_Expand_Update.TabIndex = 5
        Me.LEN_BASE_L_Expand_Update.Text = "Update"
        Me.LEN_BASE_L_Expand_Update.UseVisualStyleBackColor = True
        '
        'LEN_BASE_L_Expand_Text
        '
        Me.LEN_BASE_L_Expand_Text.Location = New System.Drawing.Point(185, 45)
        Me.LEN_BASE_L_Expand_Text.Name = "LEN_BASE_L_Expand_Text"
        Me.LEN_BASE_L_Expand_Text.Size = New System.Drawing.Size(91, 21)
        Me.LEN_BASE_L_Expand_Text.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(249, 25)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(27, 18)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "mm"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'NOW_BASE_L
        '
        Me.NOW_BASE_L.Location = New System.Drawing.Point(185, 25)
        Me.NOW_BASE_L.Name = "NOW_BASE_L"
        Me.NOW_BASE_L.Size = New System.Drawing.Size(58, 18)
        Me.NOW_BASE_L.TabIndex = 1
        Me.NOW_BASE_L.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(11, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(168, 12)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "BASE FRAME 길이(L)  늘림 :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(168, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "현재 BASE FRAME 길이(L)  :"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Auto_Trap_Option2)
        Me.GroupBox4.Controls.Add(Me.Auto_Trap_Option1)
        Me.GroupBox4.Location = New System.Drawing.Point(6, 20)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(350, 56)
        Me.GroupBox4.TabIndex = 2
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Auto Trap 사양"
        '
        'Auto_Trap_Option2
        '
        Me.Auto_Trap_Option2.AutoSize = True
        Me.Auto_Trap_Option2.Location = New System.Drawing.Point(121, 25)
        Me.Auto_Trap_Option2.Name = "Auto_Trap_Option2"
        Me.Auto_Trap_Option2.Size = New System.Drawing.Size(103, 16)
        Me.Auto_Trap_Option2.TabIndex = 0
        Me.Auto_Trap_Option2.Text = "미적용(50mm)"
        Me.Auto_Trap_Option2.UseVisualStyleBackColor = True
        '
        'Auto_Trap_Option1
        '
        Me.Auto_Trap_Option1.AutoSize = True
        Me.Auto_Trap_Option1.Checked = True
        Me.Auto_Trap_Option1.Location = New System.Drawing.Point(13, 25)
        Me.Auto_Trap_Option1.Name = "Auto_Trap_Option1"
        Me.Auto_Trap_Option1.Size = New System.Drawing.Size(97, 16)
        Me.Auto_Trap_Option1.TabIndex = 0
        Me.Auto_Trap_Option1.TabStop = True
        Me.Auto_Trap_Option1.Text = "적용(380mm)"
        Me.Auto_Trap_Option1.UseVisualStyleBackColor = True
        '
        'Frm_BaseFrame_2
        '
        Me.Frm_BaseFrame_2.Controls.Add(Me.Base_Frame_Template_Create2)
        Me.Frm_BaseFrame_2.Controls.Add(Me.DataGridView1)
        Me.Frm_BaseFrame_2.Location = New System.Drawing.Point(394, 198)
        Me.Frm_BaseFrame_2.Name = "Frm_BaseFrame_2"
        Me.Frm_BaseFrame_2.Size = New System.Drawing.Size(362, 208)
        Me.Frm_BaseFrame_2.TabIndex = 4
        Me.Frm_BaseFrame_2.TabStop = False
        '
        'Base_Frame_Template_Create2
        '
        Me.Base_Frame_Template_Create2.Location = New System.Drawing.Point(226, 169)
        Me.Base_Frame_Template_Create2.Name = "Base_Frame_Template_Create2"
        Me.Base_Frame_Template_Create2.Size = New System.Drawing.Size(130, 33)
        Me.Base_Frame_Template_Create2.TabIndex = 8
        Me.Base_Frame_Template_Create2.Text = "1. Template 실행"
        Me.Base_Frame_Template_Create2.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(6, 20)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.Size = New System.Drawing.Size(350, 143)
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
        'VVIP_Frame
        '
        Me.VVIP_Frame.Controls.Add(Me.VVIP_Message)
        Me.VVIP_Frame.Controls.Add(Me.Base_Frame_Parameter_Update)
        Me.VVIP_Frame.Controls.Add(Me.Auto_Clash_Check)
        Me.VVIP_Frame.Controls.Add(Me.UI_CONTROL_PNL)
        Me.VVIP_Frame.Location = New System.Drawing.Point(394, 12)
        Me.VVIP_Frame.Name = "VVIP_Frame"
        Me.VVIP_Frame.Size = New System.Drawing.Size(190, 586)
        Me.VVIP_Frame.TabIndex = 5
        Me.VVIP_Frame.TabStop = False
        Me.VVIP_Frame.Text = "VVIP OPTION"
        '
        'VVIP_Message
        '
        Me.VVIP_Message.Location = New System.Drawing.Point(23, 62)
        Me.VVIP_Message.Name = "VVIP_Message"
        Me.VVIP_Message.Size = New System.Drawing.Size(155, 14)
        Me.VVIP_Message.TabIndex = 7
        Me.VVIP_Message.Text = "-"
        '
        'Base_Frame_Parameter_Update
        '
        Me.Base_Frame_Parameter_Update.Location = New System.Drawing.Point(98, 79)
        Me.Base_Frame_Parameter_Update.Name = "Base_Frame_Parameter_Update"
        Me.Base_Frame_Parameter_Update.Size = New System.Drawing.Size(80, 33)
        Me.Base_Frame_Parameter_Update.TabIndex = 5
        Me.Base_Frame_Parameter_Update.Text = "Update"
        Me.Base_Frame_Parameter_Update.UseVisualStyleBackColor = True
        '
        'Auto_Clash_Check
        '
        Me.Auto_Clash_Check.AutoSize = True
        Me.Auto_Clash_Check.Checked = True
        Me.Auto_Clash_Check.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Auto_Clash_Check.Location = New System.Drawing.Point(19, 149)
        Me.Auto_Clash_Check.Name = "Auto_Clash_Check"
        Me.Auto_Clash_Check.Size = New System.Drawing.Size(128, 16)
        Me.Auto_Clash_Check.TabIndex = 4
        Me.Auto_Clash_Check.Text = "간섭체크 자동 확인"
        Me.Auto_Clash_Check.UseVisualStyleBackColor = True
        '
        'UI_CONTROL_PNL
        '
        Me.UI_CONTROL_PNL.AutoSize = True
        Me.UI_CONTROL_PNL.Location = New System.Drawing.Point(19, 40)
        Me.UI_CONTROL_PNL.Name = "UI_CONTROL_PNL"
        Me.UI_CONTROL_PNL.Size = New System.Drawing.Size(159, 16)
        Me.UI_CONTROL_PNL.TabIndex = 4
        Me.UI_CONTROL_PNL.Text = "CONTROL_PANEL_외장"
        Me.UI_CONTROL_PNL.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.ProgressBar1)
        Me.GroupBox6.Controls.Add(Me.Message_Label)
        Me.GroupBox6.Controls.Add(Me.Base_Frame_Parameter_Control)
        Me.GroupBox6.Controls.Add(Me.Base_Frame_Result_Part_Create)
        Me.GroupBox6.Controls.Add(Me.Base_Frame_Drawing_Create)
        Me.GroupBox6.Controls.Add(Me.Base_Frame_Clash_Check)
        Me.GroupBox6.Location = New System.Drawing.Point(12, 412)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(362, 186)
        Me.GroupBox6.TabIndex = 6
        Me.GroupBox6.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(10, 160)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(339, 15)
        Me.ProgressBar1.TabIndex = 7
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(8, 134)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(341, 23)
        Me.Message_Label.TabIndex = 6
        Me.Message_Label.Text = "-"
        '
        'Base_Frame_Parameter_Control
        '
        Me.Base_Frame_Parameter_Control.Location = New System.Drawing.Point(10, 20)
        Me.Base_Frame_Parameter_Control.Name = "Base_Frame_Parameter_Control"
        Me.Base_Frame_Parameter_Control.Size = New System.Drawing.Size(130, 33)
        Me.Base_Frame_Parameter_Control.TabIndex = 5
        Me.Base_Frame_Parameter_Control.Text = "Parameter Control"
        Me.Base_Frame_Parameter_Control.UseVisualStyleBackColor = True
        '
        'Base_Frame_Result_Part_Create
        '
        Me.Base_Frame_Result_Part_Create.Location = New System.Drawing.Point(219, 98)
        Me.Base_Frame_Result_Part_Create.Name = "Base_Frame_Result_Part_Create"
        Me.Base_Frame_Result_Part_Create.Size = New System.Drawing.Size(130, 33)
        Me.Base_Frame_Result_Part_Create.TabIndex = 5
        Me.Base_Frame_Result_Part_Create.Text = "3. PART 확정"
        Me.Base_Frame_Result_Part_Create.UseVisualStyleBackColor = True
        '
        'Base_Frame_Drawing_Create
        '
        Me.Base_Frame_Drawing_Create.Location = New System.Drawing.Point(219, 59)
        Me.Base_Frame_Drawing_Create.Name = "Base_Frame_Drawing_Create"
        Me.Base_Frame_Drawing_Create.Size = New System.Drawing.Size(130, 33)
        Me.Base_Frame_Drawing_Create.TabIndex = 5
        Me.Base_Frame_Drawing_Create.Text = "2. 도면 생성"
        Me.Base_Frame_Drawing_Create.UseVisualStyleBackColor = True
        '
        'Base_Frame_Clash_Check
        '
        Me.Base_Frame_Clash_Check.Location = New System.Drawing.Point(219, 20)
        Me.Base_Frame_Clash_Check.Name = "Base_Frame_Clash_Check"
        Me.Base_Frame_Clash_Check.Size = New System.Drawing.Size(130, 33)
        Me.Base_Frame_Clash_Check.TabIndex = 5
        Me.Base_Frame_Clash_Check.Text = "간섭 Check"
        Me.Base_Frame_Clash_Check.UseVisualStyleBackColor = True
        '
        'HA_Base_Frame_Template_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(385, 606)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.Frm_BaseFrame_2)
        Me.Controls.Add(Me.VVIP_Frame)
        Me.Controls.Add(Me.Frm_BaseFrame_1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "HA_Base_Frame_Template_Create"
        Me.Text = "HA_Base_Frame_Template_Create"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.Frame5.ResumeLayout(False)
        Me.Frame5.PerformLayout()
        Me.Frm_BaseFrame_1.ResumeLayout(False)
        Me.Frm_BaseFrame_1.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.Frm_BaseFrame_2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.VVIP_Frame.ResumeLayout(False)
        Me.VVIP_Frame.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
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
    Public WithEvents VVIP_Message As System.Windows.Forms.Label
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents Template_Part_Number_Text As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents UI_ENCLOSURE_TYPE_2 As System.Windows.Forms.RadioButton
    Public WithEvents UI_ENCLOSURE_TYPE As System.Windows.Forms.RadioButton
    Public WithEvents Option_Default As System.Windows.Forms.RadioButton
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents Option_Over_8 As System.Windows.Forms.RadioButton
    Public WithEvents Option_Under_8 As System.Windows.Forms.RadioButton
    Public WithEvents Option_Motor_Default As System.Windows.Forms.RadioButton
    Public WithEvents Frm_BaseFrame_1 As System.Windows.Forms.GroupBox
    Public WithEvents Base_Frame_Template_Create As System.Windows.Forms.Button
    Public WithEvents VVIP_Check As System.Windows.Forms.CheckBox
    Public WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Public WithEvents LEN_BASE_L_Expand_Update As System.Windows.Forms.Button
    Public WithEvents LEN_BASE_L_Expand_Text As System.Windows.Forms.TextBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents NOW_BASE_L As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents Auto_Trap_Option2 As System.Windows.Forms.RadioButton
    Public WithEvents Auto_Trap_Option1 As System.Windows.Forms.RadioButton
    Public WithEvents Frm_BaseFrame_2 As System.Windows.Forms.GroupBox
    Public WithEvents VVIP_Frame As System.Windows.Forms.GroupBox
    Public WithEvents Base_Frame_Parameter_Update As System.Windows.Forms.Button
    Public WithEvents UI_CONTROL_PNL As System.Windows.Forms.CheckBox
    Public WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Public WithEvents Base_Frame_Parameter_Control As System.Windows.Forms.Button
    Public WithEvents Base_Frame_Result_Part_Create As System.Windows.Forms.Button
    Public WithEvents Base_Frame_Drawing_Create As System.Windows.Forms.Button
    Public WithEvents Base_Frame_Clash_Check As System.Windows.Forms.Button
    Public WithEvents Auto_Clash_Check As System.Windows.Forms.CheckBox
    Public WithEvents Base_Frame_Template_Create2 As System.Windows.Forms.Button
End Class
