Public Class E_Oil_Piping_Create


    Private Sub E_Oil_Piping_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
        Me.Top = 150
        Me.Left = 50
        Me.Width = 601
        Me.Height = 450
        Me.TopMost = True
        Template_Data_Loading_Form.ProgressBar1_Change(20)

        '--------------------------------------------------------------------------------------------------- Coupling_Cover_Excel_Open
        If Not Oil_Piping_Excel_Open() Then
            Exit Sub
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(80)

        '--------------------------------------------------------------------------------------------------- MSFlexGrid Setting
        Call DataGridView_Inicial_Setting(DataGridView1, Oil_Piping_Selection_Parameter_Count, Oil_Piping_Selection_Parameter)

        Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)
        Result_Template_Data_Path = STW_Template_Infomation(Form_Click_Index_Value, 3)
        Result_Template_Item_Name = STW_Template_Infomation(Form_Click_Index_Value, 1)

        Template_Data_Loading_Form.ProgressBar1_Change(90)

        '--------------------------------------------------------------------------------------------------- SM기종만 Range 보이게
        If Strings.Right(sLbl_ModelNumber, 1) = "1" Then
            GBX_RangeGroupBox.Visible = False
        Else
            GBX_RangeGroupBox.Visible = True
        End If
        Call SetControlBox(Me, Me.cancel_button, True)
    End Sub


    Private Sub Oil_Piping_Create_Click(sender As Object, e As EventArgs) Handles Oil_Piping_Create.Click
        Call SetControlBox(Me, Me.cancel_button, False)
        ProgressBar1.Value = 0

        Me.Data_Path_List.Items.Clear()
        Me.DataGridView1.RowCount = 0
        Me.DataGridView1.RowCount = 1

        ProgressBar1.Value = 10
        Message_Label.Text = "Oil Piping 선정 중..."

        '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
        Call Clash_Check_Folder_Create()
        ProgressBar1.Value = 20

        '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
        Call Data_Constraint_Delete(Oil_Piping_Initial_Constraint, Oil_Piping_Initial_Constraint_Count, Result_Template_Item_Value)
        ProgressBar1.Value = 30

        '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
        If Get_Parameter_Data_Comparison(Oil_Piping_Initial_Parameter, Oil_Piping_Initial_Parameter_Count, Me) = False Then
            Exit Sub
        End If
        ProgressBar1.Value = 40

        '--------------------------------------------------------------------------------------------------- Data 비교
        Dim Part_Search_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\*.CATPart"
        Dim Part_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\"

        '--------------------------------------------------------------------------------------------------- Oil Piping Range 범위 검색
        If Range_Check.Checked = False Or Range.Text = "" Or Range.Text = "0" Then
            Call Data_Parameter_Comparison(Part_Search_Folder, Part_Folder, Me, Result_Template_Item_Value, Oil_Piping_Selection_Parameter, _
                                           Oil_Piping_Initial_Parameter, Oil_Piping_Selection_Parameter_Count)
        Else
            ' 돌아가는지, 불필요한부분 제거 필요 / lsm / 190520

            Call Data_Parameter_Comparison_Oil_Piping(Part_Search_Folder, Part_Folder, Me, Result_Template_Item_Value, Oil_Piping_Selection_Parameter, _
                                                      Oil_Piping_Initial_Parameter, Oil_Piping_Selection_Parameter_Count)
        End If
        ProgressBar1.Value = 60

        If Result_Check_Count > 1 Then
            '--------------------------------------------------------------------------------------------------- 구속조건 생성
            For i = 1 To Oil_Piping_Selection_Constraint_Count - 1
                Call Constraint_Delete(Oil_Piping_Selection_Constraint(i, 1), Oil_Piping_Selection_Constraint(i, 2), Oil_Piping_Selection_Constraint(i, 3))
            Next

            '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            Dim file_name_org = Split(Me.DataGridView1.Rows(0).Cells(2).Value, ".CATPart")
            ProgressBar1.Value = 80

            Call O_BOM_Excel_Writing(Result_Template_Item_Value, file_name_org(0))

            If Btn_Refresh_Click() = False Then
                Exit Sub
            End If

            '--------------------------------------------------------------------------------------------------- 종료 버튼 비활성화
        End If

        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        '--------------------------------------------------------------------------------------------------- 간섭체크 실행 유무 판단
        If DataGridView1.RowCount >= 1 Then
            If Not (DataGridView1.Rows(0).Cells(1).Value.ToString = "") Or (DataGridView1.Rows(0).Cells(2).Value.ToString = "") Then
                '--------------------------------------------------------------------------------------------------- 폼 X 버튼 비활성화
                'Call REMOVED_CLOSE_BUTTON(Me.Handle)
            End If
            Call SetControlBox(Me, Me.cancel_button, False)
        Else
            Call SetControlBox(Me, Me.cancel_button, True)
        End If

        ProgressBar1.Value = 100
        Message_Label.Text = "Oil Piping  선정 완료"
    End Sub


    Private Sub Oil_Piping_Clash_Check_Click(sender As Object, e As EventArgs) Handles Oil_Piping_Clash_Check.Click
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Message_Label.Text = "Oil Piping 간섭 체크중.."
        Call SetControlBox(Me, Me.cancel_button, True)
        Call Clash_Check_Folder_Create()

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change(COMP_Product_File_Name & ".CATProduct", 0, 0)

        '--------------------------------------------------------------------------------------------------- Excel Value 가져오기
        Call Get_Clash_Check_Excel_Value(Result_Template_Item_Value)

        '--------------------------------------------------------------------------------------------------- Count Check
        If Check_Count = 0 Then Exit Sub

        '--------------------------------------------------------------------------------------------------- Clash 계산 및 결과값 입력
        Call Clash_Check_Execute(Result_Template_Item_Value)
        Message_Label.Text = "Oil Piping 간섭체크 완료"
    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Call SetControlBox(Me, Me.cancel_button, False)

        Call RE_SELECTION_PART(Me, Oil_Piping_Selection_Constraint_Count, Oil_Piping_Selection_Constraint, Coupling_Cover_Initial_Constraint, Coupling_Cover_Initial_Constraint_Count, Result_Template_Item_Value)
        'Call Data_Constraint_Delete(Coupling_Cover_Initial_Constraint, Coupling_Cover_Initial_Constraint_Count, Result_Template_Item_Value)
    End Sub


    Private Sub cancel_button_Click(sender As Object, e As EventArgs) Handles cancel_button.Click
        Me.Close()
    End Sub


    Private Sub VVIP_Check_CheckedChanged(sender As Object, e As EventArgs) Handles VVIP_Check.CheckedChanged
        Call VVIP_Check_Form_Size(601, 790, Me)
    End Sub

End Class