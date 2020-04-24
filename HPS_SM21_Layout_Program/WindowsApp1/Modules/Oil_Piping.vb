Module Oil_Piping
    '--------------------------------------------------------------------------------------------------- 커플링 선택 초기화 Constraint
    Public Oil_Piping_Initial_Constraint, Oil_Piping_Initial_Constraint_Count
    '--------------------------------------------------------------------------------------------------- 커플링 선택 초기화 Parameter
    Public Oil_Piping_Initial_Parameter, Oil_Piping_Initial_Parameter_Count
    '--------------------------------------------------------------------------------------------------- 커플링 선택 Constraint
    Public Oil_Piping_Selection_Constraint, Oil_Piping_Selection_Constraint_Count
    '--------------------------------------------------------------------------------------------------- 커플링 선택 Parameter
    Public Oil_Piping_Selection_Parameter, Oil_Piping_Selection_Parameter_Count
    '--------------------------------------------------------------------------------------------------- Oil Piping Tempate Folder 저장 이름
    Public Oil_Piping_folder_name
    '--------------------------------------------------------------------------------------------------- Oil Piping Drawing File
    Public oil_piping_drawing_open
    '--------------------------------------------------------------------------------------------------- Oil Piping Bending_Type
    Public Bending_Type
    '--------------------------------------------------------------------------------------------------- Filter 관련 Parameter
    Public filter_side_direction_move, filter_move_v_length, filter_move_h_length
    '---------------------------------------------------------------------------------------------------  도면 적용 관련 Parameter
    Public EPN06_D02, EPN06_D03, EPN06_D04, Filter_Type_Value


    Public Function Oil_Piping_Excel_Open()
        sExcelFilePath = sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 1)

        '--------------------------------------------------------------------------------------------------- Oil_Piping Constraint
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(10, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          Oil_Piping_Initial_Constraint_Count, Oil_Piping_Initial_Constraint, "0")

            Template_Data_Loading_Form.ProgressBar1_Change(30)

            '--------------------------------------------------------------------------------------------------- Oil_Piping Parameter
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(10, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          Oil_Piping_Initial_Parameter_Count, Oil_Piping_Initial_Parameter, "4")

            Template_Data_Loading_Form.ProgressBar1_Change(40)

            '--------------------------------------------------------------------------------------------------- Oil_Piping  Constraint
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(11, 1), 3, "B", "C", "F", "", "", "", "", "", "", "", _
                                          Oil_Piping_Selection_Constraint_Count, Oil_Piping_Selection_Constraint, "0")

            Template_Data_Loading_Form.ProgressBar1_Change(50)

            '--------------------------------------------------------------------------------------------------- Oil_Piping Parameter
            Call Get_Excel_Value_Function(sExcelFilePath, STW_COMP_Infomation(11, 1), 10, "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", _
                                          Oil_Piping_Selection_Parameter_Count, Oil_Piping_Selection_Parameter, "4")

            Template_Data_Loading_Form.ProgressBar1_Change(60)
        Else : Return False
        End If

        '--------------------------------------------------------------------------------------------------- 도면 관련 정보 가져오기
        sExcelFilePath = Application.StartupPath & "\STW_Drawing_Infomation.xlsx"
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, assy_value_Name, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "", _
                                          STW_Drawing_Infomation_Count, STW_Drawing_Infomation, "5")
        Else : Return False
        End If

        '--------------------------------------------------------------------------------------------------- Version 정보 가져오기
        sExcelFilePath = sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4) & "\DT_Design_Rule_Version.xlsx"
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            ReDim Data_Version_Infomation(0 To 1)
            Call Get_Excel_Value_Function(sExcelFilePath, "Sheet1", 1, "A", "", "", "", "", "", "", "", "", "", _
                                          Data_Version_Infomation_Count, Data_Version_Infomation, "0")
        Else : Return False
        End If

        Template_Data_Loading_Form.ProgressBar1_Change(70)

        Return True
    End Function


    Public Sub Oil_Piping_TCV_Constraint()
        Try
            Call CATIA_BASIC02()

            constraints1 = CATIA.ActiveDocument.Product.Connections("CATIAConstraints")

            '--------------------------------------------------------------------------------------------------- TCV와 OIL COOLER 거리를 재기위한 축조립
            If assy_value_Name = "TMX" Then
                '--------------------------------------------------------------------------------------------------- TCV Constraint Delete
                Call Constraint_Delete("EPN0602_P03", "EPN060110_P01", 1)             'EPN 찾기.(유지)
                Threading.Thread.Sleep(300)

                '--------------------------------------------------------------------------------------------------- TCV Constraint Create
                Call Constraint_Delete("EPN0602_P03", "EPN060110_P01", 0)             'EPN 찾기.(유지)
                Threading.Thread.Sleep(300)
            Else
                '--------------------------------------------------------------------------------------------------- TCV Constraint Delete
                Call Constraint_Delete("EPN0602_P02", "EPN060110_P01", 1)             'EPN 찾기.(유지)

                Threading.Thread.Sleep(300)
            End If

            '--------------------------------------------------------------------------------------------------- oil_piping File Delete
            selection.Clear()

            For i = 1 To constraints1.Count
                For j = 1 To Oil_Piping_Initial_Constraint_Count - 1
                    If constraints1.Item(i).Name = Oil_Piping_Initial_Constraint(j, 2) Then
                        selection.Add(constraints1.Item(i))
                    End If
                Next
            Next

            If Not selection.Count = 0 Then
                selection.Delete()
                selection.Clear()
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Oil_Piping_TCV_Constraint() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Public Function Oil_Piping_Filter_Get_Parameter()
        Try
            Call CATIA_BASIC02()
            selection.Clear()

            '--------------------------------------------------------------------------------------------------- MOVE_FILTER_DIR_Y
            selection.search("Name='MOVE_FILTER_DIR_Y',all")
            If Not selection.Count = 0 Then
                filter_side_direction_move = selection.Item(1).Value
                EA_Oil_Piping_Template_Create.filter_side_direction_move_text.Text = filter_side_direction_move.Value
            End If
            selection.clear
            '--------------------------------------------------------------------------------------------------- MOVE_FILTER_DIR_Z
            selection.search("Name='MOVE_FILTER_DIR_Z',all")
            If Not selection.Count = 0 Then
                filter_move_v_length = selection.Item(1).Value
                EA_Oil_Piping_Template_Create.filter_move_v_length_text.Text = filter_move_v_length.Value
            End If
            selection.clear
            '--------------------------------------------------------------------------------------------------- MOVE_FILTER_DIR_X
            selection.search("Name='MOVE_FILTER_DIR_X',all")
            If Not selection.Count = 0 Then
                filter_move_h_length = selection.Item(1).Value
                EA_Oil_Piping_Template_Create.filter_move_h_length_text.Text = filter_move_h_length.Value
            End If
            selection.clear
            '    '--------------------------------------------------------------------------------------------------- 도면 Partlist 기록
            selection.search("Name='EPN060109_D01',all")             'FILTER TYPE (PAR_S, PAR_D)
            If Not selection.Count = 0 Then
                Filter_Type_Value = selection.Item(1).Value
            End If

            selection.Clear()

            Return True
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Oil_Piping_Filter_Get_Parameter() 함수 오류!!", "", 1, 1)
            Return False
        End Try
    End Function


    Public Sub Oil_Piping_Hose_Optimization()
        Try
            Dim length_value_count = 1
            Dim Check = True

            '--------------------------------------------------------------------------------------------------- Template optimization
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(Oil_Piping_folder_name & "\OIL_PIPING_Hose_Length.xls")
            Exc.Worksheets("Sheet1").Select()

            Dim S_Length_Value = Nothing
            Dim D_Length_Value = Nothing
            Do
                If Exc.Worksheets("Sheet1").Range("A" & length_value_count + 1).Value = assy_value_Name Then
                    '--------------------------------------------------------------------------------------------------- S_Length_Value , D_Length_Value 가져오기
                    S_Length_Value = Exc.Worksheets("Sheet1").Range("B" & length_value_count + 1).Value
                    D_Length_Value = Exc.Worksheets("Sheet1").Range("C" & length_value_count + 1).Value

                    '--------------------------------------------------------------------------------------------------- ac_outlet_datum_axis 유무 판별
                    Check = False
                    Exit Do
                End If
                If Exc.Worksheets("Sheet1").Range("A" & length_value_count + 1).Value = "" Then     ' 조건이 True이면
                    Check = False                                                                   ' 플래그 값을 False로 설정합니다.
                    Exit Do                                                                         ' 내부 루프를 종료합니다.
                End If
                length_value_count = length_value_count + 1
            Loop Until Check = False   ' 외부 루프를 즉시 종료합니다.

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(True)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL")

            CATIA.ActiveDocument.Product.Update()
            EA_Oil_Piping_Template_Create.ProgressBar1.Value = 90

            Call oil_piping_optimization("S_Hose_Opti_Pt", "S_Hose Length", S_Length_Value)
            Call oil_piping_optimization("D_Hose_Opti_Pt", "D_Hose Length", D_Length_Value)
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Oil_Piping_Hose_Optimization() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Public Sub oil_piping_optimization(free_parameter, target_parameter, target_value)
        Dim target_parameter_length(200) As Integer
        Dim free_parameter_value = Nothing
        Dim Target_Parameter_Value = Nothing

        '--------------------------------------------------------------------------------------------------- free_parameter , target_parameter Search
        selection.Clear()
        selection.search("Name='" & free_parameter & "',all")
        If Not selection.Count = 0 Then
            free_parameter_value = selection.Item(1).Value
            free_parameter_value.Value = 0
        End If

        selection.Clear()
        selection.search("Name='" & target_parameter & "',all")
        If Not selection.Count = 0 Then
            Target_Parameter_Value = selection.Item(1).Value
        End If

        '--------------------------------------------------------------------------------------------------- optimization_part Setting
        Dim optimization_part = Target_Parameter_Value.Parent.Parent

        Dim step_value = 0, Minus_out = 0, Plus_out = 0
        optimization_part.UpdateObject(Target_Parameter_Value)

        For i = 1 To 300
            '--------------------------------------------------------------------------------------------------- 초기 이동 위치 결정
            target_parameter_length(i) = Target_Parameter_Value.Value

            If target_parameter_length(i) = target_value Then
                Exit For
            End If

            Dim plus_parameter_value = Nothing
            Dim Minus_parameter_value = Nothing
            If i = 1 Then
                free_parameter_value.Value = free_parameter_value.Value + 0.01
                optimization_part.UpdateObject(Target_Parameter_Value)

                plus_parameter_value = Target_Parameter_Value.Value
                free_parameter_value.Value = free_parameter_value.Value - 0.02
                optimization_part.UpdateObject(Target_Parameter_Value)

                Minus_parameter_value = Target_Parameter_Value.Value
                free_parameter_value.Value = free_parameter_value.Value + 0.01
                optimization_part.UpdateObject(Target_Parameter_Value)
            End If

            '--------------------------------------------------------------------------------------------------- 목표값이 클때
            If target_value > target_parameter_length(i) Then
                '--------------------------------------------------------------------------------------------------- + 방향이 값이 커질때
                If plus_parameter_value > Minus_parameter_value Then
                    If Minus_out = 0 Then
                        free_parameter_value.Value = free_parameter_value.Value + 0.01
                    ElseIf Minus_out = 1 Then
                        free_parameter_value.Value = free_parameter_value.Value + 0.001
                        Plus_out = 1
                    ElseIf Minus_out = 2 Then
                        free_parameter_value.Value = free_parameter_value.Value + 0.0001
                        Plus_out = 2
                    ElseIf Minus_out = 3 Then
                        free_parameter_value.Value = free_parameter_value.Value + 0.00001
                        Plus_out = 3
                    End If
                Else
                    If Plus_out = 0 Then
                        free_parameter_value.Value = free_parameter_value.Value - 0.01
                        Minus_out = 1
                    ElseIf Plus_out = 1 Then
                        free_parameter_value.Value = free_parameter_value.Value - 0.001
                        Minus_out = 2
                    ElseIf Plus_out = 2 Then
                        free_parameter_value.Value = free_parameter_value.Value - 0.0001
                        Minus_out = 3
                    ElseIf Plus_out = 3 Then
                        free_parameter_value.Value = free_parameter_value.Value - 0.00001
                        Minus_out = 4
                    End If
                End If
                '--------------------------------------------------------------------------------------------------- 목표값이 작을때
            ElseIf target_value < target_parameter_length(i) Then
                '--------------------------------------------------------------------------------------------------- - 방향이 값이 작아질때
                If plus_parameter_value > Minus_parameter_value Then
                    If Plus_out = 0 Then
                        free_parameter_value.Value = free_parameter_value.Value - 0.01
                        Minus_out = 1
                    ElseIf Plus_out = 1 Then
                        free_parameter_value.Value = free_parameter_value.Value - 0.001
                        Minus_out = 2
                    ElseIf Plus_out = 2 Then
                        free_parameter_value.Value = free_parameter_value.Value - 0.0001
                        Minus_out = 3
                    ElseIf Plus_out = 3 Then
                        free_parameter_value.Value = free_parameter_value.Value - 0.00001
                        Minus_out = 4
                    End If
                Else
                    If Minus_out = 0 Then
                        free_parameter_value.Value = free_parameter_value.Value + 0.01
                    ElseIf Minus_out = 1 Then
                        free_parameter_value.Value = free_parameter_value.Value + 0.001
                        Plus_out = 1
                    ElseIf Minus_out = 2 Then
                        free_parameter_value.Value = free_parameter_value.Value + 0.0001
                        Plus_out = 2
                    ElseIf Minus_out = 3 Then
                        free_parameter_value.Value = free_parameter_value.Value + 0.00001
                        Plus_out = 3
                    End If
                End If
            End If
            optimization_part.UpdateObject(Target_Parameter_Value)

            '--------------------------------------------------------------------------------------------------- 조건 만족
            If target_parameter_length(i) - 5 >= target_value Then
                If target_parameter_length(i) + 5 <= target_value Then
                    Exit For
                End If
            End If

            If free_parameter_value.Value > 0.5 Then
                free_parameter_value.Value = 0
                optimization_part.UpdateObject(Target_Parameter_Value)
                Exit For
            ElseIf free_parameter_value.Value < -0.5 Then
                free_parameter_value.Value = 0
                optimization_part.UpdateObject(Target_Parameter_Value)
                Exit For
            End If
        Next

        optimization_part.Update()
    End Sub


    Public Function Oil_Piping_Drawing_Template_Setting()
        On Error Resume Next
        Call CATIA_BASIC03()

        '--------------------------------------------------------------------------------------------------- TABLE Detail View 가져오기
        Dim note_detail_sheet = oil_piping_drawing_open.sheets.Item("Oil_Filter (Detail)")
        Dim MyComponentRef = Nothing

        If Filter_Type_Value.Value = "INT_S" Then
            MyComponentRef = note_detail_sheet.views.Item("INT_S")
        ElseIf Filter_Type_Value.Value = "INT_D" Then
            MyComponentRef = note_detail_sheet.views.Item("INT_D")
        ElseIf Filter_Type_Value.Value = "PAR_S" Then
            MyComponentRef = note_detail_sheet.views.Item("PAR_S")
        ElseIf Filter_Type_Value.Value = "PAR_D" Then
            MyComponentRef = note_detail_sheet.views.Item("PAR_D")
        End If

        Main_View = oil_piping_drawing_open.sheets.Item("PartList").views.Item(1)

        MyComponentInst = Main_View.Components.Add(MyComponentRef, 0, 0)
        MyComponentInst.scale2 = 0.067
        DrawingSheets.Item(1).Activate()
        EA_Oil_Piping_Template_Create.ProgressBar1.Value = 50

        If Strings.Left(assy_value_Name, 2) = "SM" Then
            '---------------------------------------------------------------------------------------------------
            ' 생 성 일  : -
            ' 생 성 자  : vb6에서 사용 된 함수
            ' 수 정 일  : 19.08.08
            ' 수 정 자  : 허혜원
            ' 수정사유  : 2012년 이후 미사용
            '           : 터미네이트 에러 발생 예상 지점
            ' Parameter : -
            '---------------------------------------------------------------------------------------------------
            'Drawingselection.Clear()
            'If Bending_Type.Value = "STANDARD" Then
            '    Drawingselection.Add(oil_piping_drawing_open.sheets.Item("Limit_Bending"))
            '    Drawingselection.Add(oil_piping_drawing_open.sheets.Item("Center_Bending"))
            'ElseIf Bending_Type.Value = "LIMIT" Then
            '    Drawingselection.Add(oil_piping_drawing_open.sheets.Item("Standard_Bending"))
            '    Drawingselection.Add(oil_piping_drawing_open.sheets.Item("Center_Bending"))
            'ElseIf Bending_Type.Value = "CENTER" Then
            '    Drawingselection.Add(oil_piping_drawing_open.sheets.Item("Standard_Bending"))
            '    Drawingselection.Add(oil_piping_drawing_open.sheets.Item("Limit_Bending"))
            'End If
            'Drawingselection.Delete()
        End If

        EA_Oil_Piping_Template_Create.ProgressBar1.Value = 60

        Return True
        On Error GoTo 0
    End Function



End Module
