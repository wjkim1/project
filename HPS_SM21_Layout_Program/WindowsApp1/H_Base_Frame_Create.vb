Public Class H_Base_Frame_Create


    Private Sub H_Base_Frame_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
        Me.Top = 150
        Me.Left = 50
        Me.Width = 414
        Me.Height = 630
        Me.TopMost = True

        '--------------------------------------------------------------------------------------------------- 비활성화
        If Not Strings.Left(A_Main_Form.Lbl_ModelNumber.Text, 2) = "SM" Or Mid(A_Main_Form.Lbl_ModelNumber.Text, 4, 1) = "1" Then
            Auto_Trap_Option1.Enabled = False
            Auto_Trap_Option2.Enabled = False
            VVIP_Check.Enabled = False
        End If

        '--------------------------------------------------------------------------------------------------- Base_Frame_Excel_Open
        Call Base_Frame_Excel_Open()

        '--------------------------------------------------------------------------------------------------- Motor 무게 활성화 설정
        Frame4.Enabled = False
        Option_Motor_Default.Enabled = False
        Option_Under_8.Enabled = False
        Option_Over_8.Enabled = False

        If A_Main_Form.Lbl_ModelNumber.Text = "SM61" Then
            Frame4.Enabled = True
            Option_Motor_Default.Enabled = True
            Option_Under_8.Enabled = True
            Option_Over_8.Enabled = True
        ElseIf Strings.Right(A_Main_Form.Lbl_ModelNumber.Text, 1) = "1" Then
            '--------------------------------------------------------------------------------------------------- MSFlexGrid_Value No , Name 정의
            DataGridView1.ColumnCount = 4
            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            DataGridView1.Columns(0).HeaderText = ""
            DataGridView1.Columns(1).HeaderText = "NO"
            DataGridView1.Columns(2).HeaderText = "파일명"
            DataGridView1.Columns(3).HeaderText = "압축기 외형 크기"
            DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(0).Width = 20
            DataGridView1.Columns(1).Width = 40
            DataGridView1.Columns(2).Width = 180
            DataGridView1.Columns(2).Width = 230
        Else
            '--------------------------------------------------------------------------------------------------- MSFlexGrid Setting
            Call DataGridView_Inicial_Setting(DataGridView1, Base_Frame_Selection_Parameter_Count, Base_Frame_Selection_Parameter)
        End If
        'DataGridView1.ColumnCount = 6

        Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)
        Result_Template_Data_Path = STW_Template_Infomation(Form_Click_Index_Value, 3)
        Result_Template_Item_Name = STW_Template_Infomation(Form_Click_Index_Value, 1)

        Call SetControlBox(Me, Me.cancel_button, True)
    End Sub


    Private Sub Base_Frame_Create_Click(sender As Object, e As EventArgs) Handles Base_Frame_Create.Click
        Call SetControlBox(Me, Me.cancel_button, False)
        Try
            ProgressBar1.Value = 0
            Message_Label.Text = "Base Frame 선정 중..."

            DataGridView1.RowCount = 0
            'DataGridView1.RowCount = 1

            '--------------------------------------------------------------------------------------------------- Motor 무게
            If Frame4.Enabled = True Then
                If Option_Motor_Default.Checked = True Then
                    Call SHOW_MESSAGE_BOX("Motor 무게를 선택하세요.", "", 1, 1)
                    Exit Sub
                End If
            End If

            '--------------------------------------------------------------------------------------------------- Enclosure 유무
            If Option_Default.Checked = True Then
                Call SHOW_MESSAGE_BOX("Enclosure 유무를 선택하세요.", "", 1, 1)
                Exit Sub
            End If

            If Option_Default_MotorStarter.Checked = True Then
                Call SHOW_MESSAGE_BOX("Motor Starter 유무를 선택하세요.", "", 1, 1)
                Exit Sub
            End If

            ProgressBar1.Value = 10

            '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
            Call Clash_Check_Folder_Create()
            ProgressBar1.Value = 20

            '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
            Call Data_Constraint_Delete(Base_Frame_Initial_Constraint, Base_Frame_Initial_Constraint_Count, Result_Template_Item_Value)
            ProgressBar1.Value = 30

            If Strings.Right(A_Main_Form.Lbl_ModelNumber.Text, 1) = "1" Then
                '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
                'Call Get_Parameter_Data_Comparison(Base_Frame_Initial_Parameter, Base_Frame_Initial_Parameter_Count, Me)
                '--------------------------------------------------------------------------------------------------- Data 비교
                Dim Part_Search_Folder = A_Main_Form.Directory_Path_Text.Text & "\" & Result_Template_Data_Path & "\*.CATPart"
                Dim Part_Folder = A_Main_Form.Directory_Path_Text.Text & "\" & Result_Template_Data_Path & "\"
            End If
            ProgressBar1.Value = 40

            '--------------------------------------------------------------------------------------------------- Base 비교 Excel 값 가져오기
            Call Base_Frame_Comparison_Get_Excel()
            ProgressBar1.Value = 60

            '--------------------------------------------------------------------------------------------------- Base Frame Data Open 
            Call Base_Frame_Data_Open()
            ProgressBar1.Value = 80

            '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            If Not DataGridView1.Rows(1).Cells(2).Value = "" Then
                Dim file_name_org = Split(DataGridView1.Rows(1).Cells(2).Value, ".CATPart")

                Call O_BOM_Excel_Writing(Result_Template_Item_Value, file_name_org(0))
                Call Btn_Refresh_Click()
                Call SetControlBox(Me, Me.cancel_button, False)
            Else
                Call SetControlBox(Me, Me.cancel_button, True)
            End If
            If DataGridView1.RowCount < 1 Then
                Call SetControlBox(Me, Me.cancel_button, True)
            End If
            ProgressBar1.Value = 100
            Message_Label.Text = "Base Frame 선정 완료"

            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()
        Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Base_Frame_Create_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Base_Frame_Clash_Check_Click(sender As Object, e As EventArgs) Handles Base_Frame_Clash_Check.Click
        Message_Label.Text = "BASE FRAME 간섭 체크중.."
        Call SetControlBox(Me, Me.cancel_button, True)
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Call Clash_Check_Folder_Create()

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change(COMP_Product_File_Name & ".CATProduct", 0, 0)

        '--------------------------------------------------------------------------------------------------- Excel Value 가져오기
        Call Get_Clash_Check_Excel_Value(Result_Template_Item_Value)

        '--------------------------------------------------------------------------------------------------- Count Check
        If Check_Count = 0 Then Exit Sub

        '--------------------------------------------------------------------------------------------------- Clash 계산 및 결과값 입력
        Call Clash_Check_Execute(Result_Template_Item_Value)
        Message_Label.Text = "BASE FRAME 간섭체크 완료"
    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'cancel_button.Enabled = False
        'Call RE_SELECTION_PART(Me, Base_Frame_Selection_Constraint_Count, Base_Frame_Selection_Constraint, Base_Frame_Initial_Constraint, Base_Frame_Initial_Constraint_Count, Result_Template_Item_Value)
    End Sub


    Private Sub VVIP_Check_CheckedChanged(sender As Object, e As EventArgs) Handles VVIP_Check.CheckedChanged
        Call VVIP_Check_Form_Size(414, 632, Me)
    End Sub


    Private Sub Option_Motor_Default_CheckedChanged(sender As Object, e As EventArgs) Handles Option_Motor_Default.CheckedChanged, Option_Under_8.CheckedChanged, Option_Over_8.CheckedChanged
        If Option_Motor_Default.Checked = True Then
            Base_Frame_Create.Enabled = True
        ElseIf Option_Under_8.Checked = True Then
            Base_Frame_Create.Enabled = True
        ElseIf Option_Over_8.Checked = True Then
            Base_Frame_Create.Enabled = False
        End If
    End Sub


    Private Sub cancel_button_Click(sender As Object, e As EventArgs) Handles cancel_button.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        'DataGridView1_CellContentClick
        Call SetControlBox(Me, Me.cancel_button, False)
        Call RE_SELECTION_PART(Me, Base_Frame_Selection_Constraint_Count, Base_Frame_Selection_Constraint, Base_Frame_Initial_Constraint, Base_Frame_Initial_Constraint_Count, Result_Template_Item_Value)
    End Sub
End Class