Module Coupling_Cover
    Public Coupling_Cover_Initial_Constraint, Coupling_Cover_Initial_Constraint_Count           ' 커플링 커버 선택 초기화 Constraint
    Public Coupling_Cover_Initial_Parameter, Coupling_Cover_Initial_Parameter_Count             ' 커플링 커버 선택 초기화 Parameter
    Public Coupling_Cover_Selection_Constraint, Coupling_Cover_Selection_Constraint_Count       ' 커플링 커버 선택 Constraint
    Public Coupling_Cover_Selection_Parameter, Coupling_Cover_Selection_Parameter_Count         ' 커플링 커버 선택 Parameter
    Public coupling_cover_folder_name                                                           ' Coupling_Cover Tempate Folder 저장 이름


    Public Function Coupling_Cover_Excel_Open()
        sExcelFilePath = sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 1)

        '--------------------------------------------------------------------------------------------------- 커플링 커버 선택초기화 Constraint
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(5, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          Coupling_Cover_Initial_Constraint_Count, Coupling_Cover_Initial_Constraint, "0")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(40)

        '--------------------------------------------------------------------------------------------------- 커플링 커버 선택초기화 Parameter
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(5, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          Coupling_Cover_Initial_Parameter_Count, Coupling_Cover_Initial_Parameter, "4")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(50)

        '--------------------------------------------------------------------------------------------------- 커플링 커버 선택  Constraint
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(6, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          Coupling_Cover_Selection_Constraint_Count, Coupling_Cover_Selection_Constraint, "0")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(60)

        '--------------------------------------------------------------------------------------------------- 커플링 커버 선택 Parameter
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(6, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          Coupling_Cover_Selection_Parameter_Count, Coupling_Cover_Selection_Parameter, "4")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(70)

        '--------------------------------------------------------------------------------------------------- Version 정보 가져오기
        sExcelFilePath = sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4) & "\DT_Design_Rule_Version.xlsx"
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            ReDim Data_Version_Infomation(0 To 1)
            Call Get_Excel_Value_Function(sExcelFilePath, "Sheet1", 1, "A", "", "", "", "", "", "", "", "", "", _
                                          Data_Version_Infomation_Count, Data_Version_Infomation, "0")
        Else : Return False
        End If

        Return True
    End Function


    Public Sub Coupling_Cover_Template_Form_Inicial_Setting()
        '--------------------------------------------------------------------------------------------------- Stage_Combo 초기값 입력
        CA_Coupling_Cover_Template_Create.UI_STAGE_COMBO.Items.Add("2")
        CA_Coupling_Cover_Template_Create.UI_STAGE_COMBO.Items.Add("3")
        CA_Coupling_Cover_Template_Create.UI_STAGE_COMBO.Items.Add("4")

        '--------------------------------------------------------------------------------------------------- Update Botton Enable
        CA_Coupling_Cover_Template_Create.Coupling_Cover_Parameter_Update.Enabled = False
    End Sub

End Module
