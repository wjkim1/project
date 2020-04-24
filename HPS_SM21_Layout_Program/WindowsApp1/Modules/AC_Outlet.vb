Module AC_Outlet
    '--------------------------------------------------------------------------------------------------- AC Outlet 선택 초기화 Constraint
    Public AC_Outlet_Initial_Constraint, AC_Outlet_Initial_Constraint_Count
    '--------------------------------------------------------------------------------------------------- AC Outlet 선택 초기화 Parameter
    Public AC_Outlet_Initial_Parameter, AC_Outlet_Initial_Parameter_Count
    '--------------------------------------------------------------------------------------------------- AC Outlet 선택 Constraint
    Public AC_Outlet_Selection_Constraint, AC_Outlet_Selection_Constraint_Count
    '--------------------------------------------------------------------------------------------------- AC Outlet 선택 Parameter
    Public AC_Outlet_Selection_Parameter, AC_Outlet_Selection_Parameter_Count
    '--------------------------------------------------------------------------------------------------- AC Outlet  Tempate Folder 저장 이름
    Public AC_Outlet_folder_name
    '--------------------------------------------------------------------------------------------------- AC Outlet data replace
    Public Ac_Outlet_Open, Ac_Outlet_Data_Replace_Value(10), Ac_Outlet_Data_Replace_Name(10), Use_Block_Valve_Value
    Public Ac_Outlet_EPN_No(10), Ac_Outlet_Part_Name(10), Ac_Outlet_Part_Name_Part_No(1000)
    '--------------------------------------------------------------------------------------------------- AC Outlet Count , BOV_Rotate
    Public Count_Rotate_Parameter_Value, BOV_Rotate_Parameter_Value


    Public Function AC_Outlet_Excel_Open()
        sExcelFilePath = sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 1)

        '--------------------------------------------------------------------------------------------------- AC Outlet 선택초기화 Constraint
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(12, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          AC_Outlet_Initial_Constraint_Count, AC_Outlet_Initial_Constraint, "0")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(40)

        '--------------------------------------------------------------------------------------------------- AC Outlet 선택초기화 Parameter
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(12, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          AC_Outlet_Initial_Parameter_Count, AC_Outlet_Initial_Parameter, "4")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(50)

        '--------------------------------------------------------------------------------------------------- AC Outlet 선택 Constraint
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(13, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          AC_Outlet_Selection_Constraint_Count, AC_Outlet_Selection_Constraint, "0")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(60)

        '--------------------------------------------------------------------------------------------------- AC Outlet 선택 Parameter
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(13, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          AC_Outlet_Selection_Parameter_Count, AC_Outlet_Selection_Parameter, "4")
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


    '---------------------------------------------------------------------------------------------------
    ' Ac_Outlet_Data_Replace_Excel &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Function Ac_Outlet_Data_Replace_Excel()
        '--------------------------------------------------------------------------------------------------- Replace 대상 찾기
        Exc = CreateObject("excel.application")
        Exc.Workbooks.Open(sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 1))
        'Exc.Visible = True

        For i = 1 To 4
            Ac_Outlet_Data_Replace_Value(i) = Exc.Worksheets("AC OUTLET PIPE TEMPLATE 초기화").Range("M" & i + 1).Value
            Ac_Outlet_Data_Replace_Name(i) = Exc.Worksheets("AC OUTLET PIPE TEMPLATE 초기화").Range("N" & i + 1).Value
        Next

        Call KillProcess("EXCEL")
        '--------------------------------------------------------------------------------------------------- Replace Data와 STD_BOM 비교(PYO)
        '--------------------------------------------------------------------------------------------------- MSFlexGrid1에 Value 넣기
        FA_AC_Outlet_Template_Create.DataGridView1.RowCount = 0
        FA_AC_Outlet_Template_Create.DataGridView1.ColumnCount = 5
        FA_AC_Outlet_Template_Create.DataGridView1.Item(0, 0).Value = "No"
        FA_AC_Outlet_Template_Create.DataGridView1.Item(0, 1).Value = "EPN NO."
        FA_AC_Outlet_Template_Create.DataGridView1.Item(0, 2).Value = "O-BOM"
        FA_AC_Outlet_Template_Create.DataGridView1.Item(0, 3).Value = "DATA"
        FA_AC_Outlet_Template_Create.DataGridView1.Item(0, 4).Value = "표준품 품명"
        '--------------------------------------------------------------------------------------------------- FA_AC_Outlet_Template_Create.MSFlexGrid1 Width
        FA_AC_Outlet_Template_Create.DataGridView1.Columns(0).Width = 30
        FA_AC_Outlet_Template_Create.DataGridView1.Columns(1).Width = 120
        FA_AC_Outlet_Template_Create.DataGridView1.Columns(2).Width = 70
        FA_AC_Outlet_Template_Create.DataGridView1.Columns(3).Width = 60
        FA_AC_Outlet_Template_Create.DataGridView1.Columns(4).Width = 300
        FA_AC_Outlet_Template_Create.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'FA_AC_Outlet_Template_Create.DataGridView1.ColAlignment(1) = 1
        '--------------------------------------------------------------------------------------------------- List 생성
        For i = 1 To 4
            FA_AC_Outlet_Template_Create.DataGridView1.RowCount = i
            FA_AC_Outlet_Template_Create.DataGridView1.Item(i - 1, 0).Value = i
            FA_AC_Outlet_Template_Create.DataGridView1.Item(i - 1, 1).Value = Ac_Outlet_Data_Replace_Value(i)
            FA_AC_Outlet_Template_Create.DataGridView1.Item(i - 1, 4).Value = Ac_Outlet_Data_Replace_Name(i)
        Next

        Dim iJNum As Integer
        Dim iMNum As Integer
        '--------------------------------------------------------------------------------------------------- O-BOM 비교비교(PYO)
        For i = 1 To 4
            For j = 1 To O_BOM_Part_Value_Count - 1
                If Ac_Outlet_Data_Replace_Value(i) = O_BOM_Part_Value(j, 3) Then
                    Ac_Outlet_EPN_No(i) = O_BOM_Part_Value(j, 3)
                    Ac_Outlet_Part_Name(i) = O_BOM_Part_Value(j, 1)

                    iJNum = j
                    Exit For
                End If
            Next
            If iJNum = O_BOM_Part_Value_Count Then
                FA_AC_Outlet_Template_Create.DataGridView1.Item(i - 1, 3).Value = "V"
            End If
        Next

        '--------------------------------------------------------------------------------------------------- Data Missing 생성
        Dim xLen = Nothing
        For i = 1 To 4
            For j = 1 To O_BOM_Part_Value_Count
                If Ac_Outlet_Data_Replace_Value(i) = O_BOM_Part_Value(j, 3) Then
                    A_Main_Form.File_numbering.Path = sDirectory_Path_Text & "\" & assy_value_Name & "\" & O_BOM_Part_Value(j, 3)
                    If A_Main_Form.File_numbering.Path = sDirectory_Path_Text & "\" & assy_value_Name & "\" & O_BOM_Part_Value(j, 3) Then
                        If A_Main_Form.File_numbering.Items.Count = 0 Then
                            FA_AC_Outlet_Template_Create.DataGridView1.Item(i - 1, 3).Value = "V"
                            Exit For
                        Else
                            For m = 0 To A_Main_Form.File_numbering.Items.Count
                                xLen = Len(O_BOM_Part_Value(j, 1))
                                If Left(O_BOM_Part_Value(j, 1), xLen) = Left(A_Main_Form.File_numbering.Items(m), xLen) Then
                                    iMNum = m
                                    Exit For
                                End If
                            Next
                        End If
                        If Left(O_BOM_Part_Value(j, 1), xLen) = Left(A_Main_Form.File_numbering.Items(iMNum), xLen) Then
                            Exit For
                        End If
                        If Right(A_Main_Form.File_numbering.Items(iMNum), 7) = "CATPart" Then
                            If Not O_BOM_Part_Value(j, 1) & ".CATPart" = A_Main_Form.File_numbering.Items(iMNum) Then
                                FA_AC_Outlet_Template_Create.DataGridView1.Item(i - 1, 3).Value = "V"
                                Exit For
                            End If
                        Else
                            If Not O_BOM_Part_Value(j, 1) & ".CATProduct" = A_Main_Form.File_numbering.Items(iMNum) Then
                                FA_AC_Outlet_Template_Create.DataGridView1.Item(i - 1, 3).Value = "V"
                                Exit For
                            End If
                        End If
                    Else
                        FA_AC_Outlet_Template_Create.DataGridView1.Item(i - 1, 3).Value = "V"
                        Exit For
                    End If
                End If
            Next

            If Not Ac_Outlet_Data_Replace_Value(i) = O_BOM_Part_Value(iJNum, 3) Then
                FA_AC_Outlet_Template_Create.DataGridView1.Item(i - 1, 3).Value = "V"
            End If
        Next
        On Error GoTo 0
        '--------------------------------------------------------------------------------------------------- Block Value 유무 확인
        Dim iINum As Integer
        For i = 1 To O_BOM_Part_Value_Count
            If O_BOM_Part_Value(i, 3) = Ac_Outlet_Data_Replace_Value(4) Then
                Use_Block_Valve_Value = 1
                iINum = i
                Exit For
            End If
        Next
        If Not O_BOM_Part_Value(iINum, 3) = Ac_Outlet_Data_Replace_Value(4) Then
            Use_Block_Valve_Value = 0
        End If
    End Function


    '---------------------------------------------------------------------------------------------------
    ' Ac_Outlet_Data_Replace           &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Ac_Outlet_Data_Replace()
        Ac_Outlet_Open = CATIA.Documents.Open(AC_Outlet_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5))

        Call CATIA_BASIC02()

        '--------------------------------------------------------------------------------------------------- 기종정보 변경
        selection.Clear()
        selection.search("Name='Machine_Type',all")
        If Not selection.Count = 0 Then
            selection.Item(1).Value.Value = STW_Drawing_Infomation(1, 1)
        End If

        Dim xLen, replace_product, data_replace
        If Not Ac_Outlet_Part_Name(3) = Nothing Then
            '--------------------------------------------------------------------------------------------------- O-BOM 값이 V 가 아닐때...(PYO)
            If Not FA_AC_Outlet_Template_Create.DataGridView1.Item(3, 2).Value = "V" Then
                '--------------------------------------------------------------------------------------------------- Data 값이 V 가 아닐때...
                If Not FA_AC_Outlet_Template_Create.DataGridView1.Item(3, 3).Value = "V" Then
                    '--------------------------------------------------------------------------------------------------- BOV Data Replace
                    A_Main_Form.File_numbering.Path = sDirectory_Path_Text & "\" & assy_value_Name & "\" & Ac_Outlet_EPN_No(3)
                    If Not A_Main_Form.File_numbering.Items.Count = 0 Then
                        For m = 0 To A_Main_Form.File_numbering.Items.Count
                            xLen = Len(Ac_Outlet_Part_Name(3))
                            If Left(Ac_Outlet_Part_Name(3), xLen) = Left(A_Main_Form.File_numbering.Items(m), xLen) Then
                                replace_product = Ac_Outlet_Open.Product.products.Item(3)
                                data_replace = Ac_Outlet_Open.Product.products.ReplaceComponent(replace_product, sDirectory_Path_Text & "\" & assy_value_Name & "\" & Ac_Outlet_EPN_No(3) & "\" & A_Main_Form.File_numbering.Items(m), False)
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        End If
        '--------------------------------------------------------------------------------------------------- BOV Data Replace
        If Not Ac_Outlet_Part_Name(4) = Nothing Then
            '--------------------------------------------------------------------------------------------------- O-BOM 값이 V 가 아닐때...(pyo)
            If Not FA_AC_Outlet_Template_Create.DataGridView1.Item(4, 2).Value = "V" Then
                '--------------------------------------------------------------------------------------------------- Data 값이 V 가 아닐때...
                If Not FA_AC_Outlet_Template_Create.DataGridView1.Item(4, 3).Value = "V" Then
                    '--------------------------------------------------------------------------------------------------- BLOCK BALVE Data Replace
                    A_Main_Form.File_numbering.Path = sDirectory_Path_Text & "\" & assy_value_Name & "\" & Ac_Outlet_EPN_No(4)
                    If Not A_Main_Form.File_numbering.Items.Count = 0 Then
                        For m = 0 To A_Main_Form.File_numbering.Items.Count
                            xLen = Len(Ac_Outlet_Part_Name(4))
                            If Left(Ac_Outlet_Part_Name(4), xLen) = Left(A_Main_Form.File_numbering.Items(m), xLen) Then
                                replace_product = Ac_Outlet_Open.Product.products.Item(5)
                                data_replace = Ac_Outlet_Open.Product.products.ReplaceComponent(replace_product, sDirectory_Path_Text & "\" & assy_value_Name & "\" & Ac_Outlet_EPN_No(4) & "\" & A_Main_Form.File_numbering.Items(m), False)
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        End If
        If Not Ac_Outlet_Part_Name(1) = Nothing Then
            '--------------------------------------------------------------------------------------------------- O-BOM 값이 V 가 아닐때...
            If Not FA_AC_Outlet_Template_Create.DataGridView1.Item(1, 2).Value = "V" Then
                '--------------------------------------------------------------------------------------------------- Data 값이 V 가 아닐때...
                If Not FA_AC_Outlet_Template_Create.DataGridView1.Item(1, 3).Value = "V" Then
                    '--------------------------------------------------------------------------------------------------- DISCHANGE EXP JOINT Data Replace
                    A_Main_Form.File_numbering.Path = sDirectory_Path_Text & "\" & assy_value_Name & "\" & Ac_Outlet_EPN_No(1)
                    If Not A_Main_Form.File_numbering.Items.Count = 0 Then
                        For m = 0 To A_Main_Form.File_numbering.Items.Count
                            xLen = Len(Ac_Outlet_Part_Name(1))
                            If Left(Ac_Outlet_Part_Name(1), xLen) = Left(A_Main_Form.File_numbering.Items(m), xLen) Then
                                replace_product = Ac_Outlet_Open.Product.products.Item(6)
                                data_replace = Ac_Outlet_Open.Product.products.ReplaceComponent(replace_product, sDirectory_Path_Text & "\" & assy_value_Name & "\" & Ac_Outlet_EPN_No(1) & "\" & A_Main_Form.File_numbering.Items(m), False)
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        End If
        If Not Ac_Outlet_Part_Name(2) = Nothing Then
            '--------------------------------------------------------------------------------------------------- O-BOM 값이 V 가 아닐때...
            If Not FA_AC_Outlet_Template_Create.DataGridView1.Item(2, 2).Value = "V" Then
                '--------------------------------------------------------------------------------------------------- Data 값이 V 가 아닐때...
                If Not FA_AC_Outlet_Template_Create.DataGridView1.Item(2, 3).Value = "V" Then
                    '--------------------------------------------------------------------------------------------------- CHECK VALVE Data Replace
                    A_Main_Form.File_numbering.Path = sDirectory_Path_Text & "\" & assy_value_Name & "\" & Ac_Outlet_EPN_No(2)
                    If Not A_Main_Form.File_numbering.Items.Count = 0 Then
                        For m = 0 To A_Main_Form.File_numbering.Items.Count
                            xLen = Len(Ac_Outlet_Part_Name(2))
                            If Left(Ac_Outlet_Part_Name(2), xLen) = Left(A_Main_Form.File_numbering.Items(m), xLen) Then
                                replace_product = Ac_Outlet_Open.Product.products.Item(7)
                                data_replace = Ac_Outlet_Open.Product.products.ReplaceComponent(replace_product, sDirectory_Path_Text & "\" & assy_value_Name & "\" & Ac_Outlet_EPN_No(2) & "\" & A_Main_Form.File_numbering.Items(m), False)
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        End If

        selection.Clear()
        '--------------------------------------------------------------------------------------------------- Count_Rotate Paranmeter Value 찾기
        selection.search("Name='BLV_Rotate',all")
        If Not selection.Count = 0 Then
            Count_Rotate_Parameter_Value = selection.Item(1).Value
            FA_AC_Outlet_Template_Create.Count_Rotate_Text.Text = Count_Rotate_Parameter_Value.Value
        End If
        selection.clear
        '--------------------------------------------------------------------------------------------------- BOV_Rotate Paranmeter Value 찾기
        selection.search("Name='BOV_Rotate',all")
        If Not selection.Count = 0 Then
            BOV_Rotate_Parameter_Value = selection.Item(1).Value
            FA_AC_Outlet_Template_Create.BOV_Rotate_Text.Text = BOV_Rotate_Parameter_Value.Value
        End If
        selection.clear
        '--------------------------------------------------------------------------------------------------- 2단압축기 Parameter 변경
        selection.search("Name='STAGE',all")
        If Not selection.Count = 0 Then
            selection.Item(1).Value.Value = FA_AC_Outlet_Template_Create.Stage_Combo.Text
        End If
        selection.clear
        selection.search("Name='TYPE',all")
        If Not selection.Count = 0 Then
            selection.Item(1).Value.Value = FA_AC_Outlet_Template_Create.Type_Combo.Text
        End If
        selection.clear
        '--------------------------------------------------------------------------------------------------- Silencer 유무 Parameter 변경
        selection.search("Name='SILENCER_EXISTENCE',all")
        If Not selection.Count = 0 Then
            If FA_AC_Outlet_Template_Create.AC_Option10.Checked = True Then
                selection.Item(1).Value.Value = "true"
            ElseIf FA_AC_Outlet_Template_Create.AC_Option7.Checked = True Then
                selection.Item(1).Value.Value = "false"
            End If
        End If
        selection.clear
        selection.search("Name='BLOCK_VALVE_EXISTENCE',all")
        If Not selection.Count = 0 Then
            Dim BLOCK_VALVE_EXISTENCE = selection.Item(1).Value
            If Use_Block_Valve_Value = 1 Then
                BLOCK_VALVE_EXISTENCE.Value = True
            ElseIf Use_Block_Valve_Value = 0 Then
                BLOCK_VALVE_EXISTENCE.Value = False
            End If
        End If

        '--------------------------------------------------------------------------------------------------- Template Update
        CATIA.ActiveDocument.Product.Update()
        CATIA.ActiveDocument.Product.Update()
    End Sub


End Module
