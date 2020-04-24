Public Class G_Enclosure_Create


    Private Sub G_Enclosure_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
        Me.Top = 150
        Me.Left = 50
        Me.Width = 473
        Me.Height = 451
        Me.TopMost = True
        'cancel_button.Enabled = False

        Template_Data_Loading_Form.ProgressBar1_Change(20)

        '--------------------------------------------------------------------------------------------------- Enclosure_Excel_Open
        If Not Enclosure_Excel_Open() Then
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- MSFlexGrid Setting
        Call DataGridView_Inicial_Setting(DataGridView1, Enclosure_Selection_Parameter_Count, Enclosure_Selection_Parameter)

        '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
        Result_Template_Item_Value = "EPN1101"
        Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)
        Result_Template_Data_Path = STW_Template_Infomation(Form_Click_Index_Value, 3)
        Result_Template_Item_Name = STW_Template_Infomation(Form_Click_Index_Value, 1)

        '--------------------------------------------------------------------------------------------------- VVIP 활성화
        If Mid(A_Main_Form.Lbl_ModelNumber.Text, 4, 1) = "1" Then
            UI_EPN11_D13.Enabled = False
            UI_EPN11_D14.Enabled = False
        End If
        Call SetControlBox(Me, Me.cancel_button, True)
    End Sub


    Private Sub Enclosure_Create_Click(sender As Object, e As EventArgs) Handles Enclosure_Create.Click
        Call SetControlBox(Me, Me.cancel_button, False)
        ProgressBar1.Value = 0

        Me.Data_Path_List.Items.Clear()
        Me.DataGridView1.RowCount = 0
        Me.DataGridView1.RowCount = 1

        ProgressBar1.Value = 10
        Message_Label.Text = "Enclosure 선정 중..."

        '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
        Call Clash_Check_Folder_Create()
        ProgressBar1.Value = 20

        '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
        Call Data_Constraint_Delete(Enclosure_Initial_Constraint, Enclosure_Initial_Constraint_Count, Result_Template_Item_Value)
        ProgressBar1.Value = 30

        '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
        If Get_Parameter_Data_Comparison(Enclosure_Initial_Parameter, Enclosure_Initial_Parameter_Count, Me) = False Then
            Exit Sub
        End If
        ProgressBar1.Value = 40

        '--------------------------------------------------------------------------------------------------- Data 비교
        Dim Part_Search_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\*.CATPart"
        Dim Part_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\"

        Call Data_Parameter_Comparison(Part_Search_Folder, Part_Folder, Me, Result_Template_Item_Value, Enclosure_Selection_Parameter, Enclosure_Initial_Parameter, Enclosure_Selection_Parameter_Count)
        ProgressBar1.Value = 60

        If Result_Check_Count > 1 Then
            '--------------------------------------------------------------------------------------------------- 구속조건 생성
            For i = 1 To Enclosure_Selection_Constraint_Count - 1
                Call Constraint_Delete(Enclosure_Selection_Constraint(i, 1), Enclosure_Selection_Constraint(i, 2), Enclosure_Selection_Constraint(i, 3))
            Next

            '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            Dim file_name_org = Split(Me.DataGridView1.Rows(0).Cells(2).Value, ".CATPart")
            ProgressBar1.Value = 80

            Call O_BOM_Excel_Writing(Result_Template_Item_Value, file_name_org(0))

            If Btn_Refresh_Click() = False Then
                Exit Sub
            End If

            '--------------------------------------------------------------------------------------------------- 종료 버튼 비활성화
            Call SetControlBox(Me, Me.cancel_button, True)
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
        Message_Label.Text = "Enclosure 선정 완료"
    End Sub


    Private Sub Enclosure_Clash_Check_Click(sender As Object, e As EventArgs) Handles Enclosure_Clash_Check.Click
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Message_Label.Text = "Enclosure 간섭 체크중..."
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
        Message_Label.Text = "Enclosure 간섭체크 완료"
    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'cancel_button.Enabled = False
        Call SetControlBox(Me, Me.cancel_button, False)

        Call RE_SELECTION_PART(Me, Enclosure_Selection_Constraint_Count, Enclosure_Selection_Constraint, Enclosure_Initial_Constraint, Enclosure_Initial_Constraint_Count, Result_Template_Item_Value)
    End Sub


    Private Sub VVIP_Check_CheckedChanged(sender As Object, e As EventArgs) Handles VVIP_Check.CheckedChanged
        Call VVIP_Check_Form_Size(473, 694, Me)
    End Sub


    Private Sub cancel_button_Click(sender As Object, e As EventArgs) Handles cancel_button.Click
        Me.Close()
    End Sub

End Class