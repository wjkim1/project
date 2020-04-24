Public Class F_AC_Outlet_Create


    Private Sub F_AC_Outlet_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
        Me.Top = 150
        Me.Left = 50
        Me.Width = 558 '714
        Me.Height = 422
        Me.TopMost = True



        '--------------------------------------------------------------------------------------------------- AC_Outlet_Excel_Open
        Call AC_Outlet_Excel_Open()
        Template_Data_Loading_Form.ProgressBar1_Change(50)

        '--------------------------------------------------------------------------------------------------- MSFlexGrid Setting
        Call DataGridView_Inicial_Setting(DataGridView1, AC_Outlet_Selection_Parameter_Count, AC_Outlet_Selection_Parameter)

        Template_Data_Loading_Form.ProgressBar1_Change(75)

        Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)
        Result_Template_Data_Path = STW_Template_Infomation(Form_Click_Index_Value, 3)
        Result_Template_Item_Name = STW_Template_Infomation(Form_Click_Index_Value, 1)
        
        '--------------------------------------------------------------------------------------------------- Version 정보 가져오기
        ReDim Data_Version_Infomation(0 To 1)
        Call Get_Excel_Value_Function(sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4) & _
                                      "\DT_Design_Rule_Version.xls", "Sheet1", 1, "A", "", "", "", "", "", "", "", "", "", _
                                      Data_Version_Infomation_Count, Data_Version_Infomation, "0")

        Template_Data_Loading_Form.ProgressBar1_Change(90)
        Call SetControlBox(Me, Me.cancel_button, True)
    End Sub


    Private Sub AC_Outlet_Create_Click(sender As Object, e As EventArgs) Handles AC_Outlet_Create.Click
        Call SetControlBox(Me, Me.cancel_button, False)
        ProgressBar1.Value = 0

        DataGridView1.RowCount = 0
        DataGridView1.RowCount = 1

        ProgressBar1.Value = 10
        Message_Label.Text = "AC Outlet 선정 중..."

        '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
        Call Clash_Check_Folder_Create()
        ProgressBar1.Value = 20

        '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
        Call Data_Constraint_Delete(AC_Outlet_Initial_Constraint, AC_Outlet_Initial_Constraint_Count, Result_Template_Item_Value)
        ProgressBar1.Value = 30

        '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
        Call Get_Parameter_Data_Comparison(AC_Outlet_Initial_Parameter, AC_Outlet_Initial_Parameter_Count, Me)
        ProgressBar1.Value = 40

        '--------------------------------------------------------------------------------------------------- Data 비교
        Dim Part_Search_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\*.CATPart"
        Dim Part_Folder = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\"

        Call Data_Parameter_Comparison(Part_Search_Folder, Part_Folder, Me, Result_Template_Item_Value, AC_Outlet_Selection_Parameter,
                                       AC_Outlet_Initial_Parameter, AC_Outlet_Selection_Parameter_Count)
        ProgressBar1.Value = 60

        If Result_Check_Count > 1 Then
            '--------------------------------------------------------------------------------------------------- 구속조건 생성
            For i = 1 To AC_Outlet_Selection_Constraint_Count - 1
                Call Constraint_Delete(AC_Outlet_Selection_Constraint(i, 1), AC_Outlet_Selection_Constraint(i, 2), AC_Outlet_Selection_Constraint(i, 3))
            Next
            '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            ProgressBar1.Value = 80
            Dim file_name_org = Split(DataGridView1.Rows(0).Cells(2).Value, ".CATPart")

            Call O_BOM_Excel_Writing(Result_Template_Item_Value, file_name_org(0))
            Call Btn_Refresh_Click()
            Call SetControlBox(Me, Me.cancel_button, False)
        Else
            Call SetControlBox(Me, Me.cancel_button, True)
        End If

        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        ProgressBar1.Value = 100
        Message_Label.Text = "AC Outlet 선정 완료"
    End Sub


    Private Sub AC_Outlet_Clash_Check_Click(sender As Object, e As EventArgs) Handles AC_Outlet_Clash_Check.Click
        Call SetControlBox(Me, Me.cancel_button, True)
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Message_Label.Text = "AC Outlet 간섭 체크중.."
        Call Clash_Check_Folder_Create()

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change(COMP_Product_File_Name & ".CATProduct", 0, 0)

        '--------------------------------------------------------------------------------------------------- Excel Value 가져오기
        Call Get_Clash_Check_Excel_Value(Result_Template_Item_Value)

        '--------------------------------------------------------------------------------------------------- Count Check
        If Check_Count = 0 Then Exit Sub

        '--------------------------------------------------------------------------------------------------- Clash 계산 및 결과값 입력
        Call Clash_Check_Execute(Result_Template_Item_Value)
        Message_Label.Text = "AC Outlet 간섭체크 완료"
    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Call SetControlBox(Me, Me.cancel_button, False)

        Call RE_SELECTION_PART(Me, AC_Outlet_Selection_Constraint_Count, AC_Outlet_Selection_Constraint, AC_Outlet_Initial_Constraint, AC_Outlet_Initial_Constraint_Count, Result_Template_Item_Value)
    End Sub


    Private Sub cancel_button_Click(sender As Object, e As EventArgs) Handles cancel_button.Click
        Me.Close()
    End Sub

End Class