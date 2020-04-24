<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BA_Coupling_Template_Create
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BA_Coupling_Template_Create))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Coupling_Constraint_Excel_Path_Text = New System.Windows.Forms.Label()
        Me.Coupling_Path_Save = New System.Windows.Forms.Button()
        Me.Coupling_Constraint_Excel_Path = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Create_Publication_Parameter = New System.Windows.Forms.Button()
        Me.Coupling_Parameter_Control = New System.Windows.Forms.Button()
        Me.Coupling_BOM_File = New System.Windows.Forms.Button()
        Me.Coupling_BOM_File_Name = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Coupling_Message_Label = New System.Windows.Forms.Label()
        Me.Coupling_Template_Create = New System.Windows.Forms.Button()
        Me.Job_Repeat_Check = New System.Windows.Forms.CheckBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Coupling_Constraint_Excel_Path_Text)
        Me.GroupBox1.Controls.Add(Me.Coupling_Path_Save)
        Me.GroupBox1.Controls.Add(Me.Coupling_Constraint_Excel_Path)
        Me.GroupBox1.Location = New System.Drawing.Point(0, -7)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(528, 111)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(115, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Coupling BOM 경로"
        '
        'Coupling_Constraint_Excel_Path_Text
        '
        Me.Coupling_Constraint_Excel_Path_Text.AutoSize = True
        Me.Coupling_Constraint_Excel_Path_Text.Location = New System.Drawing.Point(12, 39)
        Me.Coupling_Constraint_Excel_Path_Text.Name = "Coupling_Constraint_Excel_Path_Text"
        Me.Coupling_Constraint_Excel_Path_Text.Size = New System.Drawing.Size(11, 12)
        Me.Coupling_Constraint_Excel_Path_Text.TabIndex = 1
        Me.Coupling_Constraint_Excel_Path_Text.Text = "-"
        '
        'Coupling_Path_Save
        '
        Me.Coupling_Path_Save.Location = New System.Drawing.Point(371, 60)
        Me.Coupling_Path_Save.Name = "Coupling_Path_Save"
        Me.Coupling_Path_Save.Size = New System.Drawing.Size(140, 36)
        Me.Coupling_Path_Save.TabIndex = 0
        Me.Coupling_Path_Save.Text = "경로 저장"
        Me.Coupling_Path_Save.UseVisualStyleBackColor = True
        '
        'Coupling_Constraint_Excel_Path
        '
        Me.Coupling_Constraint_Excel_Path.Location = New System.Drawing.Point(208, 60)
        Me.Coupling_Constraint_Excel_Path.Name = "Coupling_Constraint_Excel_Path"
        Me.Coupling_Constraint_Excel_Path.Size = New System.Drawing.Size(140, 36)
        Me.Coupling_Constraint_Excel_Path.TabIndex = 0
        Me.Coupling_Constraint_Excel_Path.Text = "재설정"
        Me.Coupling_Constraint_Excel_Path.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Create_Publication_Parameter)
        Me.GroupBox2.Controls.Add(Me.Coupling_Parameter_Control)
        Me.GroupBox2.Controls.Add(Me.Coupling_BOM_File)
        Me.GroupBox2.Controls.Add(Me.Coupling_BOM_File_Name)
        Me.GroupBox2.Location = New System.Drawing.Point(0, 110)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(528, 129)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'Create_Publication_Parameter
        '
        Me.Create_Publication_Parameter.Location = New System.Drawing.Point(371, 81)
        Me.Create_Publication_Parameter.Name = "Create_Publication_Parameter"
        Me.Create_Publication_Parameter.Size = New System.Drawing.Size(140, 36)
        Me.Create_Publication_Parameter.TabIndex = 3
        Me.Create_Publication_Parameter.Text = "2. PART 확정"
        Me.Create_Publication_Parameter.UseVisualStyleBackColor = True
        '
        'Coupling_Parameter_Control
        '
        Me.Coupling_Parameter_Control.Location = New System.Drawing.Point(208, 39)
        Me.Coupling_Parameter_Control.Name = "Coupling_Parameter_Control"
        Me.Coupling_Parameter_Control.Size = New System.Drawing.Size(140, 36)
        Me.Coupling_Parameter_Control.TabIndex = 3
        Me.Coupling_Parameter_Control.Text = "Parameter Control"
        Me.Coupling_Parameter_Control.UseVisualStyleBackColor = True
        '
        'Coupling_BOM_File
        '
        Me.Coupling_BOM_File.Location = New System.Drawing.Point(371, 39)
        Me.Coupling_BOM_File.Name = "Coupling_BOM_File"
        Me.Coupling_BOM_File.Size = New System.Drawing.Size(140, 36)
        Me.Coupling_BOM_File.TabIndex = 3
        Me.Coupling_BOM_File.Text = "1. BOM 파일 선택"
        Me.Coupling_BOM_File.UseVisualStyleBackColor = True
        '
        'Coupling_BOM_File_Name
        '
        Me.Coupling_BOM_File_Name.AutoSize = True
        Me.Coupling_BOM_File_Name.Location = New System.Drawing.Point(12, 17)
        Me.Coupling_BOM_File_Name.Name = "Coupling_BOM_File_Name"
        Me.Coupling_BOM_File_Name.Size = New System.Drawing.Size(11, 12)
        Me.Coupling_BOM_File_Name.TabIndex = 2
        Me.Coupling_BOM_File_Name.Text = "-"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ProgressBar1)
        Me.GroupBox3.Controls.Add(Me.Coupling_Message_Label)
        Me.GroupBox3.Location = New System.Drawing.Point(0, 245)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(528, 57)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 32)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(516, 12)
        Me.ProgressBar1.TabIndex = 4
        '
        'Coupling_Message_Label
        '
        Me.Coupling_Message_Label.AutoSize = True
        Me.Coupling_Message_Label.Location = New System.Drawing.Point(12, 17)
        Me.Coupling_Message_Label.Name = "Coupling_Message_Label"
        Me.Coupling_Message_Label.Size = New System.Drawing.Size(11, 12)
        Me.Coupling_Message_Label.TabIndex = 3
        Me.Coupling_Message_Label.Text = "-"
        '
        'Coupling_Template_Create
        '
        Me.Coupling_Template_Create.Location = New System.Drawing.Point(558, 53)
        Me.Coupling_Template_Create.Name = "Coupling_Template_Create"
        Me.Coupling_Template_Create.Size = New System.Drawing.Size(140, 36)
        Me.Coupling_Template_Create.TabIndex = 4
        Me.Coupling_Template_Create.Text = "2. 도면 생성"
        Me.Coupling_Template_Create.UseVisualStyleBackColor = True
        '
        'Job_Repeat_Check
        '
        Me.Job_Repeat_Check.AutoSize = True
        Me.Job_Repeat_Check.Location = New System.Drawing.Point(558, 95)
        Me.Job_Repeat_Check.Name = "Job_Repeat_Check"
        Me.Job_Repeat_Check.Size = New System.Drawing.Size(76, 16)
        Me.Job_Repeat_Check.TabIndex = 5
        Me.Job_Repeat_Check.Text = "다중 작업"
        Me.Job_Repeat_Check.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'BA_Coupling_Template_Create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(250, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(696, 374)
        Me.Controls.Add(Me.Job_Repeat_Check)
        Me.Controls.Add(Me.Coupling_Template_Create)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "BA_Coupling_Template_Create"
        Me.Text = "BA_Coupling_Template_Create"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Coupling_Constraint_Excel_Path_Text As System.Windows.Forms.Label
    Public WithEvents Coupling_Path_Save As System.Windows.Forms.Button
    Public WithEvents Coupling_Constraint_Excel_Path As System.Windows.Forms.Button
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents Create_Publication_Parameter As System.Windows.Forms.Button
    Public WithEvents Coupling_Parameter_Control As System.Windows.Forms.Button
    Public WithEvents Coupling_BOM_File As System.Windows.Forms.Button
    Public WithEvents Coupling_BOM_File_Name As System.Windows.Forms.Label
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Public WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents Coupling_Message_Label As System.Windows.Forms.Label
    Public WithEvents Coupling_Template_Create As System.Windows.Forms.Button
    Public WithEvents Job_Repeat_Check As System.Windows.Forms.CheckBox
    Public WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
End Class
