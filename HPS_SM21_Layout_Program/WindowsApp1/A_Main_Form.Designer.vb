'Namespace main_form
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class A_Main_Form
        Inherits System.Windows.Forms.Form

        'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
        <System.Diagnostics.DebuggerNonUserCode()>
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
        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(A_Main_Form))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TP1_GB1 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Directory_Path_Text = New System.Windows.Forms.TextBox()
        Me.Directory_Path_Cmd = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.O_BOM_Path_Text = New System.Windows.Forms.TextBox()
        Me.O_BOM_Path_Cmd = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Constraint_Excel_Path_Text = New System.Windows.Forms.TextBox()
        Me.Constraint_Excel_Path_Cmd = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Project_DB_Path_Text = New System.Windows.Forms.TextBox()
        Me.Project_DB_Path_Cmd = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Template_Data_Path_Text = New System.Windows.Forms.TextBox()
        Me.Template_Data_Path_Cmd = New System.Windows.Forms.Button()
        Me.DB_Path_Save = New System.Windows.Forms.Button()
        Me.Lbl_ModelNumber = New System.Windows.Forms.Label()
        Me.MainTab = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Clash_Check = New System.Windows.Forms.Button()
        Me.Over_Constraint_List = New System.Windows.Forms.Button()
        Me.Element_ALL_Hide = New System.Windows.Forms.Button()
        Me.retry_constraint = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Col5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TP1_GB2 = New System.Windows.Forms.GroupBox()
        Me.O_BOM_File_Name = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Result_Data_Replace = New System.Windows.Forms.Button()
        Me.Get_Layout_List = New System.Windows.Forms.Button()
        Me.Get_Now_O_BOM_File = New System.Windows.Forms.Button()
        Me.O_BOM_File = New System.Windows.Forms.Button()
        Me.TP2 = New System.Windows.Forms.TabPage()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Title_Block_Create = New System.Windows.Forms.Button()
        Me.Oli_Piping_Re_Constraint = New System.Windows.Forms.Button()
        Me.MainTab_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Template_Create9 = New System.Windows.Forms.Button()
        Me.Result_Part_Create9 = New System.Windows.Forms.Button()
        Me.Template_Create8 = New System.Windows.Forms.Button()
        Me.Result_Part_Create5 = New System.Windows.Forms.Button()
        Me.Result_Part_Create8 = New System.Windows.Forms.Button()
        Me.Template_Create5 = New System.Windows.Forms.Button()
        Me.Template_Create4 = New System.Windows.Forms.Button()
        Me.Result_Part_Create2 = New System.Windows.Forms.Button()
        Me.Result_Part_Create4 = New System.Windows.Forms.Button()
        Me.Template_Create2 = New System.Windows.Forms.Button()
        Me.Template_Create7 = New System.Windows.Forms.Button()
        Me.Result_Part_Create6 = New System.Windows.Forms.Button()
        Me.Result_Part_Create7 = New System.Windows.Forms.Button()
        Me.Template_Create6 = New System.Windows.Forms.Button()
        Me.Template_Create3 = New System.Windows.Forms.Button()
        Me.Result_Part_Create3 = New System.Windows.Forms.Button()
        Me.Result_Part_Create1 = New System.Windows.Forms.Button()
        Me.Template_Create1 = New System.Windows.Forms.Button()
        Me.Template_Name_Label9 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Template_Name_Label8 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Template_Name_Label7 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Template_Name_Label6 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Template_Name_Label5 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Template_Name_Label4 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Template_Name_Label3 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Template_Name_Label2 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Template_Name_Label1 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.Rotor_Assy_Create = New System.Windows.Forms.Button()
        Me.Coupling_Template_Create = New System.Windows.Forms.Button()
        Me.Motor_Template_Create = New System.Windows.Forms.Button()
        Me.TP3 = New System.Windows.Forms.TabPage()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.Ref_DWG_Create = New System.Windows.Forms.Button()
        Me.Ref_Drawing = New System.Windows.Forms.Button()
        Me.Base_DWG_Create = New System.Windows.Forms.Button()
        Me.Base_Drawing = New System.Windows.Forms.Button()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.Comp_Layout_Save = New System.Windows.Forms.Button()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.Project_Data_Path_Open = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.project_data_path = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.File_numbering = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox()
        Me.re_o_bom_path = New System.Windows.Forms.Label()
        Me.data_open_list = New System.Windows.Forms.ListBox()
        Me.data_open_path_list = New System.Windows.Forms.ListBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Radio1 = New System.Windows.Forms.RadioButton()
        Me.Radio2 = New System.Windows.Forms.RadioButton()
        Me.Check1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.BtEx = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Label = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.TP1_GB1.SuspendLayout()
        Me.MainTab.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TP1_GB2.SuspendLayout()
        Me.TP2.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.MainTab_GroupBox.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.TP3.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TP1_GB1)
        Me.GroupBox1.Controls.Add(Me.MainTab)
        Me.GroupBox1.Location = New System.Drawing.Point(-1, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(790, 640)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'TP1_GB1
        '
        Me.TP1_GB1.Controls.Add(Me.Label6)
        Me.TP1_GB1.Controls.Add(Me.Label1)
        Me.TP1_GB1.Controls.Add(Me.Directory_Path_Text)
        Me.TP1_GB1.Controls.Add(Me.Directory_Path_Cmd)
        Me.TP1_GB1.Controls.Add(Me.Label3)
        Me.TP1_GB1.Controls.Add(Me.O_BOM_Path_Text)
        Me.TP1_GB1.Controls.Add(Me.O_BOM_Path_Cmd)
        Me.TP1_GB1.Controls.Add(Me.Label2)
        Me.TP1_GB1.Controls.Add(Me.Constraint_Excel_Path_Text)
        Me.TP1_GB1.Controls.Add(Me.Constraint_Excel_Path_Cmd)
        Me.TP1_GB1.Controls.Add(Me.Label4)
        Me.TP1_GB1.Controls.Add(Me.Project_DB_Path_Text)
        Me.TP1_GB1.Controls.Add(Me.Project_DB_Path_Cmd)
        Me.TP1_GB1.Controls.Add(Me.Label5)
        Me.TP1_GB1.Controls.Add(Me.Template_Data_Path_Text)
        Me.TP1_GB1.Controls.Add(Me.Template_Data_Path_Cmd)
        Me.TP1_GB1.Controls.Add(Me.DB_Path_Save)
        Me.TP1_GB1.Controls.Add(Me.Lbl_ModelNumber)
        Me.TP1_GB1.Location = New System.Drawing.Point(6, 20)
        Me.TP1_GB1.Name = "TP1_GB1"
        Me.TP1_GB1.Size = New System.Drawing.Size(778, 117)
        Me.TP1_GB1.TabIndex = 0
        Me.TP1_GB1.TabStop = False
        Me.TP1_GB1.Text = "작업중인 DB 정보"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(123, 0)
        Me.Label6.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(65, 12)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "기종 정보 :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(30, 27)
        Me.Label1.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 12)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "3D 형상 DB :"
        '
        'Directory_Path_Text
        '
        Me.Directory_Path_Text.Location = New System.Drawing.Point(108, 24)
        Me.Directory_Path_Text.Name = "Directory_Path_Text"
        Me.Directory_Path_Text.ReadOnly = True
        Me.Directory_Path_Text.Size = New System.Drawing.Size(195, 21)
        Me.Directory_Path_Text.TabIndex = 0
        '
        'Directory_Path_Cmd
        '
        Me.Directory_Path_Cmd.Location = New System.Drawing.Point(309, 22)
        Me.Directory_Path_Cmd.Name = "Directory_Path_Cmd"
        Me.Directory_Path_Cmd.Size = New System.Drawing.Size(49, 23)
        Me.Directory_Path_Cmd.TabIndex = 0
        Me.Directory_Path_Cmd.Text = "재설정"
        Me.Directory_Path_Cmd.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(29, 56)
        Me.Label3.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(76, 12)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "O-BOM DB :"
        '
        'O_BOM_Path_Text
        '
        Me.O_BOM_Path_Text.Location = New System.Drawing.Point(108, 53)
        Me.O_BOM_Path_Text.Name = "O_BOM_Path_Text"
        Me.O_BOM_Path_Text.ReadOnly = True
        Me.O_BOM_Path_Text.Size = New System.Drawing.Size(195, 21)
        Me.O_BOM_Path_Text.TabIndex = 2
        '
        'O_BOM_Path_Cmd
        '
        Me.O_BOM_Path_Cmd.Location = New System.Drawing.Point(309, 51)
        Me.O_BOM_Path_Cmd.Name = "O_BOM_Path_Cmd"
        Me.O_BOM_Path_Cmd.Size = New System.Drawing.Size(49, 23)
        Me.O_BOM_Path_Cmd.TabIndex = 2
        Me.O_BOM_Path_Cmd.Text = "재설정"
        Me.O_BOM_Path_Cmd.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(373, 27)
        Me.Label2.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 12)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "연결  정보 : "
        '
        'Constraint_Excel_Path_Text
        '
        Me.Constraint_Excel_Path_Text.Location = New System.Drawing.Point(447, 24)
        Me.Constraint_Excel_Path_Text.Name = "Constraint_Excel_Path_Text"
        Me.Constraint_Excel_Path_Text.ReadOnly = True
        Me.Constraint_Excel_Path_Text.Size = New System.Drawing.Size(195, 21)
        Me.Constraint_Excel_Path_Text.TabIndex = 1
        '
        'Constraint_Excel_Path_Cmd
        '
        Me.Constraint_Excel_Path_Cmd.Location = New System.Drawing.Point(648, 22)
        Me.Constraint_Excel_Path_Cmd.Name = "Constraint_Excel_Path_Cmd"
        Me.Constraint_Excel_Path_Cmd.Size = New System.Drawing.Size(49, 23)
        Me.Constraint_Excel_Path_Cmd.TabIndex = 1
        Me.Constraint_Excel_Path_Cmd.Text = "재설정"
        Me.Constraint_Excel_Path_Cmd.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(369, 56)
        Me.Label4.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 12)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Project DB :"
        '
        'Project_DB_Path_Text
        '
        Me.Project_DB_Path_Text.Location = New System.Drawing.Point(447, 53)
        Me.Project_DB_Path_Text.Name = "Project_DB_Path_Text"
        Me.Project_DB_Path_Text.ReadOnly = True
        Me.Project_DB_Path_Text.Size = New System.Drawing.Size(195, 21)
        Me.Project_DB_Path_Text.TabIndex = 3
        '
        'Project_DB_Path_Cmd
        '
        Me.Project_DB_Path_Cmd.Location = New System.Drawing.Point(648, 51)
        Me.Project_DB_Path_Cmd.Name = "Project_DB_Path_Cmd"
        Me.Project_DB_Path_Cmd.Size = New System.Drawing.Size(49, 23)
        Me.Project_DB_Path_Cmd.TabIndex = 3
        Me.Project_DB_Path_Cmd.Text = "재설정"
        Me.Project_DB_Path_Cmd.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 85)
        Me.Label5.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(99, 12)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "TEMPLATE DB :"
        '
        'Template_Data_Path_Text
        '
        Me.Template_Data_Path_Text.Location = New System.Drawing.Point(108, 82)
        Me.Template_Data_Path_Text.Name = "Template_Data_Path_Text"
        Me.Template_Data_Path_Text.ReadOnly = True
        Me.Template_Data_Path_Text.Size = New System.Drawing.Size(195, 21)
        Me.Template_Data_Path_Text.TabIndex = 4
        '
        'Template_Data_Path_Cmd
        '
        Me.Template_Data_Path_Cmd.Location = New System.Drawing.Point(309, 80)
        Me.Template_Data_Path_Cmd.Name = "Template_Data_Path_Cmd"
        Me.Template_Data_Path_Cmd.Size = New System.Drawing.Size(49, 23)
        Me.Template_Data_Path_Cmd.TabIndex = 4
        Me.Template_Data_Path_Cmd.Text = "재설정"
        Me.Template_Data_Path_Cmd.UseVisualStyleBackColor = True
        '
        'DB_Path_Save
        '
        Me.DB_Path_Save.Location = New System.Drawing.Point(705, 21)
        Me.DB_Path_Save.Name = "DB_Path_Save"
        Me.DB_Path_Save.Size = New System.Drawing.Size(67, 52)
        Me.DB_Path_Save.TabIndex = 5
        Me.DB_Path_Save.Text = "경로 저장"
        Me.DB_Path_Save.UseVisualStyleBackColor = True
        '
        'Lbl_ModelNumber
        '
        Me.Lbl_ModelNumber.Location = New System.Drawing.Point(194, 0)
        Me.Lbl_ModelNumber.Margin = New System.Windows.Forms.Padding(3, 7, 3, 0)
        Me.Lbl_ModelNumber.Name = "Lbl_ModelNumber"
        Me.Lbl_ModelNumber.Size = New System.Drawing.Size(72, 12)
        Me.Lbl_ModelNumber.TabIndex = 4
        Me.Lbl_ModelNumber.Text = "정보 없음"
        '
        'MainTab
        '
        Me.MainTab.Controls.Add(Me.TabPage1)
        Me.MainTab.Controls.Add(Me.TP2)
        Me.MainTab.Controls.Add(Me.TP3)
        Me.MainTab.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.MainTab.ItemSize = New System.Drawing.Size(153, 18)
        Me.MainTab.Location = New System.Drawing.Point(6, 143)
        Me.MainTab.Name = "MainTab"
        Me.MainTab.SelectedIndex = 0
        Me.MainTab.Size = New System.Drawing.Size(778, 492)
        Me.MainTab.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TabPage1.Controls.Add(Me.GroupBox3)
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Controls.Add(Me.TP1_GB2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(770, 466)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "         O-BOM 3D LAYOUT 구성         "
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Clash_Check)
        Me.GroupBox3.Controls.Add(Me.Over_Constraint_List)
        Me.GroupBox3.Controls.Add(Me.Element_ALL_Hide)
        Me.GroupBox3.Controls.Add(Me.retry_constraint)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 412)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(758, 48)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "LAYOUT 구성 부가기능"
        '
        'Clash_Check
        '
        Me.Clash_Check.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Clash_Check.Location = New System.Drawing.Point(573, 17)
        Me.Clash_Check.Name = "Clash_Check"
        Me.Clash_Check.Size = New System.Drawing.Size(174, 23)
        Me.Clash_Check.TabIndex = 3
        Me.Clash_Check.Text = "Layout 간섭 Check"
        Me.Clash_Check.UseVisualStyleBackColor = True
        '
        'Over_Constraint_List
        '
        Me.Over_Constraint_List.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Over_Constraint_List.Location = New System.Drawing.Point(11, 17)
        Me.Over_Constraint_List.Name = "Over_Constraint_List"
        Me.Over_Constraint_List.Size = New System.Drawing.Size(174, 23)
        Me.Over_Constraint_List.TabIndex = 0
        Me.Over_Constraint_List.Text = "O/Constraint List"
        Me.Over_Constraint_List.UseVisualStyleBackColor = True
        '
        'Element_ALL_Hide
        '
        Me.Element_ALL_Hide.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Element_ALL_Hide.Location = New System.Drawing.Point(199, 17)
        Me.Element_ALL_Hide.Name = "Element_ALL_Hide"
        Me.Element_ALL_Hide.Size = New System.Drawing.Size(174, 23)
        Me.Element_ALL_Hide.TabIndex = 1
        Me.Element_ALL_Hide.Text = "Geometry ALL Hide"
        Me.Element_ALL_Hide.UseVisualStyleBackColor = True
        '
        'retry_constraint
        '
        Me.retry_constraint.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.retry_constraint.Location = New System.Drawing.Point(386, 17)
        Me.retry_constraint.Name = "retry_constraint"
        Me.retry_constraint.Size = New System.Drawing.Size(174, 23)
        Me.retry_constraint.TabIndex = 2
        Me.retry_constraint.Text = "Constraint 재구성"
        Me.retry_constraint.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col1, Me.Col2, Me.Col3, Me.Col4, Me.Col5})
        Me.DataGridView1.Location = New System.Drawing.Point(6, 84)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(758, 322)
        Me.DataGridView1.TabIndex = 0
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
        'TP1_GB2
        '
        Me.TP1_GB2.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TP1_GB2.Controls.Add(Me.O_BOM_File_Name)
        Me.TP1_GB2.Controls.Add(Me.Label8)
        Me.TP1_GB2.Controls.Add(Me.Result_Data_Replace)
        Me.TP1_GB2.Controls.Add(Me.Get_Layout_List)
        Me.TP1_GB2.Controls.Add(Me.Get_Now_O_BOM_File)
        Me.TP1_GB2.Controls.Add(Me.O_BOM_File)
        Me.TP1_GB2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.TP1_GB2.Location = New System.Drawing.Point(6, 7)
        Me.TP1_GB2.Name = "TP1_GB2"
        Me.TP1_GB2.Size = New System.Drawing.Size(758, 71)
        Me.TP1_GB2.TabIndex = 0
        Me.TP1_GB2.TabStop = False
        Me.TP1_GB2.Text = " LAYOUT 구성하기 "
        '
        'O_BOM_File_Name
        '
        Me.O_BOM_File_Name.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.O_BOM_File_Name.Location = New System.Drawing.Point(116, 16)
        Me.O_BOM_File_Name.Name = "O_BOM_File_Name"
        Me.O_BOM_File_Name.Size = New System.Drawing.Size(624, 22)
        Me.O_BOM_File_Name.TabIndex = 6
        Me.O_BOM_File_Name.Text = "-"
        Me.O_BOM_File_Name.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label8.Location = New System.Drawing.Point(15, 20)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(105, 17)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "진행중 O-BOM :"
        '
        'Result_Data_Replace
        '
        Me.Result_Data_Replace.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Data_Replace.Location = New System.Drawing.Point(573, 40)
        Me.Result_Data_Replace.Name = "Result_Data_Replace"
        Me.Result_Data_Replace.Size = New System.Drawing.Size(174, 23)
        Me.Result_Data_Replace.TabIndex = 3
        Me.Result_Data_Replace.Text = "Part Replace"
        Me.Result_Data_Replace.UseVisualStyleBackColor = True
        '
        'Get_Layout_List
        '
        Me.Get_Layout_List.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Get_Layout_List.Location = New System.Drawing.Point(386, 40)
        Me.Get_Layout_List.Name = "Get_Layout_List"
        Me.Get_Layout_List.Size = New System.Drawing.Size(174, 23)
        Me.Get_Layout_List.TabIndex = 2
        Me.Get_Layout_List.Text = "진행중 Product"
        Me.Get_Layout_List.UseVisualStyleBackColor = True
        '
        'Get_Now_O_BOM_File
        '
        Me.Get_Now_O_BOM_File.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Get_Now_O_BOM_File.Location = New System.Drawing.Point(199, 40)
        Me.Get_Now_O_BOM_File.Name = "Get_Now_O_BOM_File"
        Me.Get_Now_O_BOM_File.Size = New System.Drawing.Size(174, 23)
        Me.Get_Now_O_BOM_File.TabIndex = 1
        Me.Get_Now_O_BOM_File.Text = "진행중 O-BOM"
        Me.Get_Now_O_BOM_File.UseVisualStyleBackColor = True
        '
        'O_BOM_File
        '
        Me.O_BOM_File.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.O_BOM_File.Location = New System.Drawing.Point(11, 40)
        Me.O_BOM_File.Name = "O_BOM_File"
        Me.O_BOM_File.Size = New System.Drawing.Size(174, 23)
        Me.O_BOM_File.TabIndex = 0
        Me.O_BOM_File.Text = "O-BOM File 가져오기"
        Me.O_BOM_File.UseVisualStyleBackColor = True
        '
        'TP2
        '
        Me.TP2.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TP2.Controls.Add(Me.GroupBox6)
        Me.TP2.Controls.Add(Me.MainTab_GroupBox)
        Me.TP2.Controls.Add(Me.GroupBox5)
        Me.TP2.Location = New System.Drawing.Point(4, 22)
        Me.TP2.Name = "TP2"
        Me.TP2.Padding = New System.Windows.Forms.Padding(3)
        Me.TP2.Size = New System.Drawing.Size(770, 466)
        Me.TP2.TabIndex = 1
        Me.TP2.Text = "            실적형상 적용하기             "
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Button4)
        Me.GroupBox6.Controls.Add(Me.Button3)
        Me.GroupBox6.Controls.Add(Me.Title_Block_Create)
        Me.GroupBox6.Controls.Add(Me.Oli_Piping_Re_Constraint)
        Me.GroupBox6.Location = New System.Drawing.Point(15, 403)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(734, 55)
        Me.GroupBox6.TabIndex = 17
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "부가기능"
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button4.Location = New System.Drawing.Point(555, 17)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(160, 28)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "Utility"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Enabled = False
        Me.Button3.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Button3.Location = New System.Drawing.Point(376, 18)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(160, 28)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "EPN ALL 간섭체크"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Title_Block_Create
        '
        Me.Title_Block_Create.Enabled = False
        Me.Title_Block_Create.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Title_Block_Create.Location = New System.Drawing.Point(197, 17)
        Me.Title_Block_Create.Name = "Title_Block_Create"
        Me.Title_Block_Create.Size = New System.Drawing.Size(160, 28)
        Me.Title_Block_Create.TabIndex = 1
        Me.Title_Block_Create.Text = "SIGNATURES 기입"
        Me.Title_Block_Create.UseVisualStyleBackColor = True
        '
        'Oli_Piping_Re_Constraint
        '
        Me.Oli_Piping_Re_Constraint.Enabled = False
        Me.Oli_Piping_Re_Constraint.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Oli_Piping_Re_Constraint.Location = New System.Drawing.Point(18, 18)
        Me.Oli_Piping_Re_Constraint.Name = "Oli_Piping_Re_Constraint"
        Me.Oli_Piping_Re_Constraint.Size = New System.Drawing.Size(160, 28)
        Me.Oli_Piping_Re_Constraint.TabIndex = 0
        Me.Oli_Piping_Re_Constraint.Text = "OIL PIPING 재구성"
        Me.Oli_Piping_Re_Constraint.UseVisualStyleBackColor = True
        '
        'MainTab_GroupBox
        '
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Create9)
        Me.MainTab_GroupBox.Controls.Add(Me.Result_Part_Create9)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Create8)
        Me.MainTab_GroupBox.Controls.Add(Me.Result_Part_Create5)
        Me.MainTab_GroupBox.Controls.Add(Me.Result_Part_Create8)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Create5)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Create4)
        Me.MainTab_GroupBox.Controls.Add(Me.Result_Part_Create2)
        Me.MainTab_GroupBox.Controls.Add(Me.Result_Part_Create4)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Create2)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Create7)
        Me.MainTab_GroupBox.Controls.Add(Me.Result_Part_Create6)
        Me.MainTab_GroupBox.Controls.Add(Me.Result_Part_Create7)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Create6)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Create3)
        Me.MainTab_GroupBox.Controls.Add(Me.Result_Part_Create3)
        Me.MainTab_GroupBox.Controls.Add(Me.Result_Part_Create1)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Create1)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Name_Label9)
        Me.MainTab_GroupBox.Controls.Add(Me.Label23)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Name_Label8)
        Me.MainTab_GroupBox.Controls.Add(Me.Label21)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Name_Label7)
        Me.MainTab_GroupBox.Controls.Add(Me.Label19)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Name_Label6)
        Me.MainTab_GroupBox.Controls.Add(Me.Label17)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Name_Label5)
        Me.MainTab_GroupBox.Controls.Add(Me.Label15)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Name_Label4)
        Me.MainTab_GroupBox.Controls.Add(Me.Label13)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Name_Label3)
        Me.MainTab_GroupBox.Controls.Add(Me.Label11)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Name_Label2)
        Me.MainTab_GroupBox.Controls.Add(Me.Label9)
        Me.MainTab_GroupBox.Controls.Add(Me.Template_Name_Label1)
        Me.MainTab_GroupBox.Controls.Add(Me.Label7)
        Me.MainTab_GroupBox.Location = New System.Drawing.Point(15, 6)
        Me.MainTab_GroupBox.Name = "MainTab_GroupBox"
        Me.MainTab_GroupBox.Size = New System.Drawing.Size(734, 335)
        Me.MainTab_GroupBox.TabIndex = 16
        Me.MainTab_GroupBox.TabStop = False
        '
        'Template_Create9
        '
        Me.Template_Create9.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Create9.Location = New System.Drawing.Point(517, 299)
        Me.Template_Create9.Name = "Template_Create9"
        Me.Template_Create9.Size = New System.Drawing.Size(199, 28)
        Me.Template_Create9.TabIndex = 17
        Me.Template_Create9.Text = "TEMPLATE 생성"
        Me.Template_Create9.UseVisualStyleBackColor = True
        '
        'Result_Part_Create9
        '
        Me.Result_Part_Create9.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Part_Create9.Location = New System.Drawing.Point(18, 299)
        Me.Result_Part_Create9.Name = "Result_Part_Create9"
        Me.Result_Part_Create9.Size = New System.Drawing.Size(199, 28)
        Me.Result_Part_Create9.TabIndex = 16
        Me.Result_Part_Create9.Text = "실적 선택"
        Me.Result_Part_Create9.UseVisualStyleBackColor = True
        '
        'Template_Create8
        '
        Me.Template_Create8.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Create8.Location = New System.Drawing.Point(517, 263)
        Me.Template_Create8.Name = "Template_Create8"
        Me.Template_Create8.Size = New System.Drawing.Size(199, 28)
        Me.Template_Create8.TabIndex = 15
        Me.Template_Create8.Text = "TEMPLATE 생성"
        Me.Template_Create8.UseVisualStyleBackColor = True
        '
        'Result_Part_Create5
        '
        Me.Result_Part_Create5.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Part_Create5.Location = New System.Drawing.Point(18, 155)
        Me.Result_Part_Create5.Name = "Result_Part_Create5"
        Me.Result_Part_Create5.Size = New System.Drawing.Size(199, 28)
        Me.Result_Part_Create5.TabIndex = 8
        Me.Result_Part_Create5.Text = "실적 선택"
        Me.Result_Part_Create5.UseVisualStyleBackColor = True
        '
        'Result_Part_Create8
        '
        Me.Result_Part_Create8.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Part_Create8.Location = New System.Drawing.Point(18, 263)
        Me.Result_Part_Create8.Name = "Result_Part_Create8"
        Me.Result_Part_Create8.Size = New System.Drawing.Size(199, 28)
        Me.Result_Part_Create8.TabIndex = 14
        Me.Result_Part_Create8.Text = "실적 선택"
        Me.Result_Part_Create8.UseVisualStyleBackColor = True
        '
        'Template_Create5
        '
        Me.Template_Create5.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Create5.Location = New System.Drawing.Point(517, 155)
        Me.Template_Create5.Name = "Template_Create5"
        Me.Template_Create5.Size = New System.Drawing.Size(199, 28)
        Me.Template_Create5.TabIndex = 9
        Me.Template_Create5.Text = "TEMPLATE 생성"
        Me.Template_Create5.UseVisualStyleBackColor = True
        '
        'Template_Create4
        '
        Me.Template_Create4.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Create4.Location = New System.Drawing.Point(517, 119)
        Me.Template_Create4.Name = "Template_Create4"
        Me.Template_Create4.Size = New System.Drawing.Size(199, 28)
        Me.Template_Create4.TabIndex = 7
        Me.Template_Create4.Text = "TEMPLATE 생성"
        Me.Template_Create4.UseVisualStyleBackColor = True
        '
        'Result_Part_Create2
        '
        Me.Result_Part_Create2.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Part_Create2.Location = New System.Drawing.Point(18, 47)
        Me.Result_Part_Create2.Name = "Result_Part_Create2"
        Me.Result_Part_Create2.Size = New System.Drawing.Size(199, 28)
        Me.Result_Part_Create2.TabIndex = 2
        Me.Result_Part_Create2.Text = "실적 선택"
        Me.Result_Part_Create2.UseVisualStyleBackColor = True
        '
        'Result_Part_Create4
        '
        Me.Result_Part_Create4.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Part_Create4.Location = New System.Drawing.Point(18, 119)
        Me.Result_Part_Create4.Name = "Result_Part_Create4"
        Me.Result_Part_Create4.Size = New System.Drawing.Size(199, 28)
        Me.Result_Part_Create4.TabIndex = 6
        Me.Result_Part_Create4.Text = "실적 선택"
        Me.Result_Part_Create4.UseVisualStyleBackColor = True
        '
        'Template_Create2
        '
        Me.Template_Create2.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Create2.Location = New System.Drawing.Point(517, 47)
        Me.Template_Create2.Name = "Template_Create2"
        Me.Template_Create2.Size = New System.Drawing.Size(199, 28)
        Me.Template_Create2.TabIndex = 3
        Me.Template_Create2.Text = "TEMPLATE 생성"
        Me.Template_Create2.UseVisualStyleBackColor = True
        '
        'Template_Create7
        '
        Me.Template_Create7.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Create7.Location = New System.Drawing.Point(517, 227)
        Me.Template_Create7.Name = "Template_Create7"
        Me.Template_Create7.Size = New System.Drawing.Size(199, 28)
        Me.Template_Create7.TabIndex = 13
        Me.Template_Create7.Text = "TEMPLATE 생성"
        Me.Template_Create7.UseVisualStyleBackColor = True
        '
        'Result_Part_Create6
        '
        Me.Result_Part_Create6.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Part_Create6.Location = New System.Drawing.Point(18, 191)
        Me.Result_Part_Create6.Name = "Result_Part_Create6"
        Me.Result_Part_Create6.Size = New System.Drawing.Size(199, 28)
        Me.Result_Part_Create6.TabIndex = 10
        Me.Result_Part_Create6.Text = "실적 선택"
        Me.Result_Part_Create6.UseVisualStyleBackColor = True
        '
        'Result_Part_Create7
        '
        Me.Result_Part_Create7.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Part_Create7.Location = New System.Drawing.Point(18, 227)
        Me.Result_Part_Create7.Name = "Result_Part_Create7"
        Me.Result_Part_Create7.Size = New System.Drawing.Size(199, 28)
        Me.Result_Part_Create7.TabIndex = 12
        Me.Result_Part_Create7.Text = "실적 선택"
        Me.Result_Part_Create7.UseVisualStyleBackColor = True
        '
        'Template_Create6
        '
        Me.Template_Create6.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Create6.Location = New System.Drawing.Point(517, 191)
        Me.Template_Create6.Name = "Template_Create6"
        Me.Template_Create6.Size = New System.Drawing.Size(199, 28)
        Me.Template_Create6.TabIndex = 11
        Me.Template_Create6.Text = "TEMPLATE 생성"
        Me.Template_Create6.UseVisualStyleBackColor = True
        '
        'Template_Create3
        '
        Me.Template_Create3.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Create3.Location = New System.Drawing.Point(517, 83)
        Me.Template_Create3.Name = "Template_Create3"
        Me.Template_Create3.Size = New System.Drawing.Size(199, 28)
        Me.Template_Create3.TabIndex = 5
        Me.Template_Create3.Text = "TEMPLATE 생성"
        Me.Template_Create3.UseVisualStyleBackColor = True
        '
        'Result_Part_Create3
        '
        Me.Result_Part_Create3.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Part_Create3.Location = New System.Drawing.Point(18, 83)
        Me.Result_Part_Create3.Name = "Result_Part_Create3"
        Me.Result_Part_Create3.Size = New System.Drawing.Size(199, 28)
        Me.Result_Part_Create3.TabIndex = 4
        Me.Result_Part_Create3.Text = "실적 선택"
        Me.Result_Part_Create3.UseVisualStyleBackColor = True
        '
        'Result_Part_Create1
        '
        Me.Result_Part_Create1.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Result_Part_Create1.Location = New System.Drawing.Point(18, 13)
        Me.Result_Part_Create1.Name = "Result_Part_Create1"
        Me.Result_Part_Create1.Size = New System.Drawing.Size(199, 28)
        Me.Result_Part_Create1.TabIndex = 0
        Me.Result_Part_Create1.Text = "실적 선택"
        Me.Result_Part_Create1.UseVisualStyleBackColor = True
        '
        'Template_Create1
        '
        Me.Template_Create1.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Create1.Location = New System.Drawing.Point(517, 13)
        Me.Template_Create1.Name = "Template_Create1"
        Me.Template_Create1.Size = New System.Drawing.Size(199, 28)
        Me.Template_Create1.TabIndex = 1
        Me.Template_Create1.Text = "TEMPLATE 생성"
        Me.Template_Create1.UseVisualStyleBackColor = True
        '
        'Template_Name_Label9
        '
        Me.Template_Name_Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Template_Name_Label9.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Name_Label9.Location = New System.Drawing.Point(270, 299)
        Me.Template_Name_Label9.Name = "Template_Name_Label9"
        Me.Template_Name_Label9.Size = New System.Drawing.Size(195, 27)
        Me.Template_Name_Label9.TabIndex = 2
        Me.Template_Name_Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label23
        '
        Me.Label23.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label23.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label23.Location = New System.Drawing.Point(214, 311)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(306, 2)
        Me.Label23.TabIndex = 2
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Template_Name_Label8
        '
        Me.Template_Name_Label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Template_Name_Label8.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Name_Label8.Location = New System.Drawing.Point(270, 264)
        Me.Template_Name_Label8.Name = "Template_Name_Label8"
        Me.Template_Name_Label8.Size = New System.Drawing.Size(195, 27)
        Me.Template_Name_Label8.TabIndex = 2
        Me.Template_Name_Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label21
        '
        Me.Label21.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label21.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label21.Location = New System.Drawing.Point(214, 276)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(306, 2)
        Me.Label21.TabIndex = 2
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Template_Name_Label7
        '
        Me.Template_Name_Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Template_Name_Label7.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Name_Label7.Location = New System.Drawing.Point(270, 228)
        Me.Template_Name_Label7.Name = "Template_Name_Label7"
        Me.Template_Name_Label7.Size = New System.Drawing.Size(195, 27)
        Me.Template_Name_Label7.TabIndex = 2
        Me.Template_Name_Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label19
        '
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label19.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label19.Location = New System.Drawing.Point(214, 240)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(306, 2)
        Me.Label19.TabIndex = 2
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Template_Name_Label6
        '
        Me.Template_Name_Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Template_Name_Label6.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Name_Label6.Location = New System.Drawing.Point(270, 192)
        Me.Template_Name_Label6.Name = "Template_Name_Label6"
        Me.Template_Name_Label6.Size = New System.Drawing.Size(195, 27)
        Me.Template_Name_Label6.TabIndex = 2
        Me.Template_Name_Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label17
        '
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label17.Location = New System.Drawing.Point(214, 204)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(306, 2)
        Me.Label17.TabIndex = 2
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Template_Name_Label5
        '
        Me.Template_Name_Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Template_Name_Label5.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Name_Label5.Location = New System.Drawing.Point(270, 156)
        Me.Template_Name_Label5.Name = "Template_Name_Label5"
        Me.Template_Name_Label5.Size = New System.Drawing.Size(195, 27)
        Me.Template_Name_Label5.TabIndex = 2
        Me.Template_Name_Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label15
        '
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label15.Location = New System.Drawing.Point(214, 168)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(306, 2)
        Me.Label15.TabIndex = 2
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Template_Name_Label4
        '
        Me.Template_Name_Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Template_Name_Label4.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Name_Label4.Location = New System.Drawing.Point(270, 120)
        Me.Template_Name_Label4.Name = "Template_Name_Label4"
        Me.Template_Name_Label4.Size = New System.Drawing.Size(195, 27)
        Me.Template_Name_Label4.TabIndex = 2
        Me.Template_Name_Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label13
        '
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label13.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label13.Location = New System.Drawing.Point(214, 132)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(306, 2)
        Me.Label13.TabIndex = 2
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Template_Name_Label3
        '
        Me.Template_Name_Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Template_Name_Label3.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Name_Label3.Location = New System.Drawing.Point(270, 84)
        Me.Template_Name_Label3.Name = "Template_Name_Label3"
        Me.Template_Name_Label3.Size = New System.Drawing.Size(195, 27)
        Me.Template_Name_Label3.TabIndex = 2
        Me.Template_Name_Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label11.Location = New System.Drawing.Point(214, 96)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(306, 2)
        Me.Label11.TabIndex = 2
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Template_Name_Label2
        '
        Me.Template_Name_Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Template_Name_Label2.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Name_Label2.Location = New System.Drawing.Point(270, 48)
        Me.Template_Name_Label2.Name = "Template_Name_Label2"
        Me.Template_Name_Label2.Size = New System.Drawing.Size(195, 27)
        Me.Template_Name_Label2.TabIndex = 2
        Me.Template_Name_Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label9.Location = New System.Drawing.Point(214, 60)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(306, 2)
        Me.Label9.TabIndex = 2
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Template_Name_Label1
        '
        Me.Template_Name_Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Template_Name_Label1.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Template_Name_Label1.Location = New System.Drawing.Point(270, 14)
        Me.Template_Name_Label1.Name = "Template_Name_Label1"
        Me.Template_Name_Label1.Size = New System.Drawing.Size(195, 27)
        Me.Template_Name_Label1.TabIndex = 2
        Me.Template_Name_Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label7.Location = New System.Drawing.Point(214, 26)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(306, 2)
        Me.Label7.TabIndex = 2
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Rotor_Assy_Create)
        Me.GroupBox5.Controls.Add(Me.Coupling_Template_Create)
        Me.GroupBox5.Controls.Add(Me.Motor_Template_Create)
        Me.GroupBox5.Location = New System.Drawing.Point(15, 345)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(734, 55)
        Me.GroupBox5.TabIndex = 16
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "단일 Template"
        '
        'Rotor_Assy_Create
        '
        Me.Rotor_Assy_Create.Enabled = False
        Me.Rotor_Assy_Create.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Rotor_Assy_Create.Location = New System.Drawing.Point(266, 17)
        Me.Rotor_Assy_Create.Name = "Rotor_Assy_Create"
        Me.Rotor_Assy_Create.Size = New System.Drawing.Size(199, 28)
        Me.Rotor_Assy_Create.TabIndex = 1
        Me.Rotor_Assy_Create.Text = "Rotor Assy Create 생성"
        Me.Rotor_Assy_Create.UseVisualStyleBackColor = True
        '
        'Coupling_Template_Create
        '
        Me.Coupling_Template_Create.Enabled = False
        Me.Coupling_Template_Create.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Coupling_Template_Create.Location = New System.Drawing.Point(517, 17)
        Me.Coupling_Template_Create.Name = "Coupling_Template_Create"
        Me.Coupling_Template_Create.Size = New System.Drawing.Size(199, 28)
        Me.Coupling_Template_Create.TabIndex = 2
        Me.Coupling_Template_Create.Text = "COUPLING Template 생성"
        Me.Coupling_Template_Create.UseVisualStyleBackColor = True
        '
        'Motor_Template_Create
        '
        Me.Motor_Template_Create.Enabled = False
        Me.Motor_Template_Create.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Motor_Template_Create.Location = New System.Drawing.Point(18, 17)
        Me.Motor_Template_Create.Name = "Motor_Template_Create"
        Me.Motor_Template_Create.Size = New System.Drawing.Size(199, 28)
        Me.Motor_Template_Create.TabIndex = 0
        Me.Motor_Template_Create.Text = "Motor Template 생성"
        Me.Motor_Template_Create.UseVisualStyleBackColor = True
        '
        'TP3
        '
        Me.TP3.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TP3.Controls.Add(Me.GroupBox9)
        Me.TP3.Controls.Add(Me.GroupBox8)
        Me.TP3.Controls.Add(Me.GroupBox7)
        Me.TP3.Location = New System.Drawing.Point(4, 22)
        Me.TP3.Name = "TP3"
        Me.TP3.Padding = New System.Windows.Forms.Padding(3)
        Me.TP3.Size = New System.Drawing.Size(770, 466)
        Me.TP3.TabIndex = 2
        Me.TP3.Text = "      COMP LAYOUT DRAWING 생성      "
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.Ref_DWG_Create)
        Me.GroupBox9.Controls.Add(Me.Ref_Drawing)
        Me.GroupBox9.Controls.Add(Me.Base_DWG_Create)
        Me.GroupBox9.Controls.Add(Me.Base_Drawing)
        Me.GroupBox9.Location = New System.Drawing.Point(15, 306)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(734, 140)
        Me.GroupBox9.TabIndex = 9
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "COMP LAYOUT 도면 생성"
        '
        'Ref_DWG_Create
        '
        Me.Ref_DWG_Create.Enabled = False
        Me.Ref_DWG_Create.Location = New System.Drawing.Point(239, 82)
        Me.Ref_DWG_Create.Name = "Ref_DWG_Create"
        Me.Ref_DWG_Create.Size = New System.Drawing.Size(200, 40)
        Me.Ref_DWG_Create.TabIndex = 12
        Me.Ref_DWG_Create.Text = "참고도면 DWG 생성"
        Me.Ref_DWG_Create.UseVisualStyleBackColor = True
        '
        'Ref_Drawing
        '
        Me.Ref_Drawing.Location = New System.Drawing.Point(23, 82)
        Me.Ref_Drawing.Name = "Ref_Drawing"
        Me.Ref_Drawing.Size = New System.Drawing.Size(200, 40)
        Me.Ref_Drawing.TabIndex = 13
        Me.Ref_Drawing.Text = "참고도면 생성"
        Me.Ref_Drawing.UseVisualStyleBackColor = True
        '
        'Base_DWG_Create
        '
        Me.Base_DWG_Create.Enabled = False
        Me.Base_DWG_Create.Location = New System.Drawing.Point(239, 31)
        Me.Base_DWG_Create.Name = "Base_DWG_Create"
        Me.Base_DWG_Create.Size = New System.Drawing.Size(200, 40)
        Me.Base_DWG_Create.TabIndex = 11
        Me.Base_DWG_Create.Text = "기본도면 DWG 생성"
        Me.Base_DWG_Create.UseVisualStyleBackColor = True
        '
        'Base_Drawing
        '
        Me.Base_Drawing.Enabled = False
        Me.Base_Drawing.Location = New System.Drawing.Point(23, 31)
        Me.Base_Drawing.Name = "Base_Drawing"
        Me.Base_Drawing.Size = New System.Drawing.Size(200, 40)
        Me.Base_Drawing.TabIndex = 11
        Me.Base_Drawing.Text = "기본도면 설정"
        Me.Base_Drawing.UseVisualStyleBackColor = True
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.Comp_Layout_Save)
        Me.GroupBox8.Location = New System.Drawing.Point(15, 160)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(734, 140)
        Me.GroupBox8.TabIndex = 8
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "COMP LAYOUT 확정"
        '
        'Comp_Layout_Save
        '
        Me.Comp_Layout_Save.Location = New System.Drawing.Point(23, 31)
        Me.Comp_Layout_Save.Name = "Comp_Layout_Save"
        Me.Comp_Layout_Save.Size = New System.Drawing.Size(200, 40)
        Me.Comp_Layout_Save.TabIndex = 10
        Me.Comp_Layout_Save.Text = "COMP LAYOUT 형상 확정"
        Me.Comp_Layout_Save.UseVisualStyleBackColor = True
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.Project_Data_Path_Open)
        Me.GroupBox7.Controls.Add(Me.Label10)
        Me.GroupBox7.Controls.Add(Me.project_data_path)
        Me.GroupBox7.Location = New System.Drawing.Point(15, 14)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(734, 140)
        Me.GroupBox7.TabIndex = 8
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "COMP LAYOUT 확정 경로 선택"
        '
        'Project_Data_Path_Open
        '
        Me.Project_Data_Path_Open.Location = New System.Drawing.Point(23, 82)
        Me.Project_Data_Path_Open.Name = "Project_Data_Path_Open"
        Me.Project_Data_Path_Open.Size = New System.Drawing.Size(200, 40)
        Me.Project_Data_Path_Open.TabIndex = 9
        Me.Project_Data_Path_Open.Text = "재설정"
        Me.Project_Data_Path_Open.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(229, 94)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(119, 17)
        Me.Label10.TabIndex = 8
        Me.Label10.Text = "(별도 설정시 사용)"
        '
        'project_data_path
        '
        Me.project_data_path.Font = New System.Drawing.Font("맑은 고딕", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.project_data_path.Location = New System.Drawing.Point(10, 36)
        Me.project_data_path.Name = "project_data_path"
        Me.project_data_path.Size = New System.Drawing.Size(504, 30)
        Me.project_data_path.TabIndex = 8
        Me.project_data_path.Text = "-"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.File_numbering)
        Me.GroupBox2.Controls.Add(Me.re_o_bom_path)
        Me.GroupBox2.Controls.Add(Me.data_open_list)
        Me.GroupBox2.Controls.Add(Me.data_open_path_list)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 740)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(783, 215)
        Me.GroupBox2.TabIndex = 12
        Me.GroupBox2.TabStop = False
        '
        'File_numbering
        '
        Me.File_numbering.FormattingEnabled = True
        Me.File_numbering.Location = New System.Drawing.Point(395, 17)
        Me.File_numbering.Name = "File_numbering"
        Me.File_numbering.Pattern = "*.*"
        Me.File_numbering.Size = New System.Drawing.Size(333, 148)
        Me.File_numbering.TabIndex = 14
        '
        're_o_bom_path
        '
        Me.re_o_bom_path.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.re_o_bom_path.Location = New System.Drawing.Point(6, 177)
        Me.re_o_bom_path.Name = "re_o_bom_path"
        Me.re_o_bom_path.Size = New System.Drawing.Size(570, 17)
        Me.re_o_bom_path.TabIndex = 12
        Me.re_o_bom_path.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'data_open_list
        '
        Me.data_open_list.FormattingEnabled = True
        Me.data_open_list.ItemHeight = 12
        Me.data_open_list.Location = New System.Drawing.Point(6, 99)
        Me.data_open_list.Name = "data_open_list"
        Me.data_open_list.Size = New System.Drawing.Size(378, 64)
        Me.data_open_list.TabIndex = 9
        '
        'data_open_path_list
        '
        Me.data_open_path_list.FormattingEnabled = True
        Me.data_open_path_list.ItemHeight = 12
        Me.data_open_path_list.Location = New System.Drawing.Point(6, 17)
        Me.data_open_path_list.Name = "data_open_path_list"
        Me.data_open_path_list.Size = New System.Drawing.Size(378, 76)
        Me.data_open_path_list.TabIndex = 9
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Radio1
        '
        Me.Radio1.AutoSize = True
        Me.Radio1.Location = New System.Drawing.Point(840, 91)
        Me.Radio1.Name = "Radio1"
        Me.Radio1.Size = New System.Drawing.Size(124, 16)
        Me.Radio1.TabIndex = 13
        Me.Radio1.TabStop = True
        Me.Radio1.Text = "Element ALL Hide"
        Me.Radio1.UseVisualStyleBackColor = True
        '
        'Radio2
        '
        Me.Radio2.AutoSize = True
        Me.Radio2.Location = New System.Drawing.Point(840, 120)
        Me.Radio2.Name = "Radio2"
        Me.Radio2.Size = New System.Drawing.Size(84, 16)
        Me.Radio2.TabIndex = 14
        Me.Radio2.TabStop = True
        Me.Radio2.Text = "Axis Show"
        Me.Radio2.UseVisualStyleBackColor = True
        '
        'Check1
        '
        Me.Check1.AutoSize = True
        Me.Check1.Checked = True
        Me.Check1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Check1.Location = New System.Drawing.Point(840, 60)
        Me.Check1.Name = "Check1"
        Me.Check1.Size = New System.Drawing.Size(104, 16)
        Me.Check1.TabIndex = 15
        Me.Check1.Text = "연결 과정 보기"
        Me.Check1.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.BtEx)
        Me.GroupBox4.Controls.Add(Me.ProgressBar1)
        Me.GroupBox4.Controls.Add(Me.Message_Label)
        Me.GroupBox4.Location = New System.Drawing.Point(-1, 633)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(790, 80)
        Me.GroupBox4.TabIndex = 16
        Me.GroupBox4.TabStop = False
        '
        'BtEx
        '
        Me.BtEx.Font = New System.Drawing.Font("굴림", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.BtEx.Location = New System.Drawing.Point(619, 13)
        Me.BtEx.Name = "BtEx"
        Me.BtEx.Size = New System.Drawing.Size(155, 30)
        Me.BtEx.TabIndex = 0
        Me.BtEx.Text = "종료"
        Me.BtEx.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(16, 49)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(758, 22)
        Me.ProgressBar1.TabIndex = 13
        '
        'Message_Label
        '
        Me.Message_Label.Location = New System.Drawing.Point(16, 16)
        Me.Message_Label.Name = "Message_Label"
        Me.Message_Label.Size = New System.Drawing.Size(597, 25)
        Me.Message_Label.TabIndex = 12
        Me.Message_Label.Text = "-"
        Me.Message_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'A_Main_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1030, 882)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.Check1)
        Me.Controls.Add(Me.Radio2)
        Me.Controls.Add(Me.Radio1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "A_Main_Form"
        Me.Text = "압축기 Layout Program"
        Me.GroupBox1.ResumeLayout(False)
        Me.TP1_GB1.ResumeLayout(False)
        Me.TP1_GB1.PerformLayout()
        Me.MainTab.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TP1_GB2.ResumeLayout(False)
        Me.TP1_GB2.PerformLayout()
        Me.TP2.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.MainTab_GroupBox.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.TP3.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents GroupBox1 As GroupBox
    Public WithEvents TP1_GB1 As GroupBox
    Public WithEvents DB_Path_Save As Button
    Public WithEvents Lbl_ModelNumber As Label
    Public WithEvents Label6 As Label
    Public WithEvents GroupBox2 As GroupBox
    Public WithEvents data_open_list As ListBox
    Public WithEvents data_open_path_list As ListBox
    Public WithEvents OpenFileDialog1 As OpenFileDialog
    Public WithEvents re_o_bom_path As System.Windows.Forms.Label
    Public WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents Project_DB_Path_Text As System.Windows.Forms.TextBox
    Public WithEvents Project_DB_Path_Cmd As System.Windows.Forms.Button
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Template_Data_Path_Text As System.Windows.Forms.TextBox
    Public WithEvents Template_Data_Path_Cmd As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Directory_Path_Text As System.Windows.Forms.TextBox
    Public WithEvents Directory_Path_Cmd As System.Windows.Forms.Button
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents O_BOM_Path_Text As System.Windows.Forms.TextBox
    Public WithEvents O_BOM_Path_Cmd As System.Windows.Forms.Button
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Constraint_Excel_Path_Text As System.Windows.Forms.TextBox
    Public WithEvents Constraint_Excel_Path_Cmd As System.Windows.Forms.Button
    Public WithEvents BtEx As System.Windows.Forms.Button
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Message_Label As System.Windows.Forms.Label
    Public WithEvents Radio1 As System.Windows.Forms.RadioButton
    Public WithEvents Radio2 As System.Windows.Forms.RadioButton
    Public WithEvents Check1 As System.Windows.Forms.CheckBox
    Public WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents File_numbering As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
    Public WithEvents MainTab As TabControl
    Public WithEvents TabPage1 As TabPage
    Public WithEvents GroupBox3 As GroupBox
    Public WithEvents Clash_Check As Button
    Public WithEvents Over_Constraint_List As Button
    Public WithEvents Element_ALL_Hide As Button
    Public WithEvents retry_constraint As Button
    Public WithEvents DataGridView1 As DataGridView
    Public WithEvents Col1 As DataGridViewTextBoxColumn
    Public WithEvents Col2 As DataGridViewTextBoxColumn
    Public WithEvents Col3 As DataGridViewTextBoxColumn
    Public WithEvents Col4 As DataGridViewTextBoxColumn
    Public WithEvents Col5 As DataGridViewTextBoxColumn
    Public WithEvents TP1_GB2 As GroupBox
    Public WithEvents O_BOM_File_Name As Label
    Public WithEvents Label8 As Label
    Public WithEvents Result_Data_Replace As Button
    Public WithEvents Get_Layout_List As Button
    Public WithEvents Get_Now_O_BOM_File As Button
    Public WithEvents O_BOM_File As Button
    Public WithEvents TP2 As TabPage
    Public WithEvents GroupBox6 As GroupBox
    Public WithEvents Button4 As Button
    Public WithEvents Button3 As Button
    Public WithEvents Title_Block_Create As Button
    Public WithEvents Oli_Piping_Re_Constraint As Button
    Public WithEvents MainTab_GroupBox As GroupBox
    Public WithEvents Template_Create9 As Button
    Public WithEvents Result_Part_Create9 As Button
    Public WithEvents Template_Create8 As Button
    Public WithEvents Result_Part_Create5 As Button
    Public WithEvents Result_Part_Create8 As Button
    Public WithEvents Template_Create5 As Button
    Public WithEvents Template_Create4 As Button
    Public WithEvents Result_Part_Create2 As Button
    Public WithEvents Result_Part_Create4 As Button
    Public WithEvents Template_Create2 As Button
    Public WithEvents Template_Create7 As Button
    Public WithEvents Result_Part_Create6 As Button
    Public WithEvents Result_Part_Create7 As Button
    Public WithEvents Template_Create6 As Button
    Public WithEvents Template_Create3 As Button
    Public WithEvents Result_Part_Create3 As Button
    Public WithEvents Result_Part_Create1 As Button
    Public WithEvents Template_Create1 As Button
    Public WithEvents Template_Name_Label9 As Label
    Public WithEvents Label23 As Label
    Public WithEvents Template_Name_Label8 As Label
    Public WithEvents Label21 As Label
    Public WithEvents Template_Name_Label7 As Label
    Public WithEvents Label19 As Label
    Public WithEvents Template_Name_Label6 As Label
    Public WithEvents Label17 As Label
    Public WithEvents Template_Name_Label5 As Label
    Public WithEvents Label15 As Label
    Public WithEvents Template_Name_Label4 As Label
    Public WithEvents Label13 As Label
    Public WithEvents Template_Name_Label3 As Label
    Public WithEvents Label11 As Label
    Public WithEvents Template_Name_Label2 As Label
    Public WithEvents Label9 As Label
    Public WithEvents Template_Name_Label1 As Label
    Public WithEvents Label7 As Label
    Public WithEvents GroupBox5 As GroupBox
    Public WithEvents Rotor_Assy_Create As Button
    Public WithEvents Coupling_Template_Create As Button
    Public WithEvents Motor_Template_Create As Button
    Public WithEvents TP3 As TabPage
    Friend WithEvents GroupBox9 As GroupBox
    Friend WithEvents Ref_DWG_Create As Button
    Friend WithEvents Ref_Drawing As Button
    Friend WithEvents Base_DWG_Create As Button
    Friend WithEvents Base_Drawing As Button
    Friend WithEvents GroupBox8 As GroupBox
    Friend WithEvents Comp_Layout_Save As Button
    Friend WithEvents GroupBox7 As GroupBox
    Friend WithEvents Project_Data_Path_Open As Button
    Public WithEvents Label10 As Label
    Public WithEvents project_data_path As Label
End Class

