Public Class EA_Oil_Piping_Template_Create


    Private Sub EA_Oil_Piping_Template_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
            Me.Top = 150
            Me.Left = 50
            Me.Width = 362
            Me.Height = 452
            Me.TopMost = True

            Call CATIA_BASIC02()

            Template_Data_Loading_Form.ProgressBar1_Change(30)

            '--------------------------------------------------------------------------------------------------- Oil_Piping_Template_Form_Inicial_Setting
            Oil_Piping_Parameter_Update.Enabled = False
            Oil_Piping_Drawing_Create.Enabled = False
            Oil_Piping_Part_Convert.Enabled = False

            '--------------------------------------------------------------------------------------------------- Type 활성화
            '---------------------------------------------------------------------------------------------------
            ' 생 성 일  : -
            ' 생 성 자  : 
            ' 수 정 일  : 20.04.23
            ' 수 정 자  : 김원주
            ' 수정사유  : SM2100기종은 filter 이동 기능 필요 없음.
            '             추가 -> GroupBox2.Enabled = True
            ' Parameter : -
            '---------------------------------------------------------------------------------------------------

            'If Strings.Left(assy_value_Name, 2) = "SM" Then
            '    filter_move_h_length_text.Enabled = True
            '    filter_side_direction_move_text.Enabled = True
            '    filter_move_v_length_text.Enabled = True
            '    FILTER_MOVE_H_LENGTH_Update.Enabled = True
            '    FILTER_SIDE_DIRECTION_MOVE_Update.Enabled = True
            '    FILTER_MOVE_V_LENGTH_Update.Enabled = True
            '    Label11.Enabled = True
            '    Label1.Enabled = True
            '    Label2.Enabled = True
            'Else
            '    filter_move_h_length_text.Enabled = False
            '    filter_side_direction_move_text.Enabled = False
            '    filter_move_v_length_text.Enabled = False
            '    FILTER_MOVE_H_LENGTH_Update.Enabled = False
            '    FILTER_SIDE_DIRECTION_MOVE_Update.Enabled = False
            '    FILTER_MOVE_V_LENGTH_Update.Enabled = False
            '    Label11.Enabled = False
            '    Label1.Enabled = False
            '    Label2.Enabled = False
            'End If
            GroupBox2.Enabled = True
            '--------------------------------------------------------------------------------------------------- BOM Path 및 날자 가져오기
            Call Data_O_BOM_Excel_Search()

            Template_Data_Loading_Form.ProgressBar1_Change(40)

            '--------------------------------------------------------------------------------------------------- Oil_Piping_Excel_Open
            Call Oil_Piping_Excel_Open()

            Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)
        Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("EA_Oil_Piping_Template_Create_Load() 함수 오류!!", "", 1, 1) '터미네이터
        End Try
    End Sub


    Private Sub Oil_Piping_Template_Create_Click(sender As Object, e As EventArgs) Handles Oil_Piping_Template_Create.Click
        '--------------------------------------------------------------------------------------------------- Oil Piping Part NO 가 없을때
        If Template_Part_Number_Text.Text = "" Then
            Call SHOW_MESSAGE_BOX("Oil Piping Part No.가 없습니다.", "", 1, 1)
            Exit Sub
        End If

        Try
            '--------------------------------------------------------------------------------------------------- Template 실행여부 확인
            If SHOW_MESSAGE_BOX("Template01 실행 하겠습니까?", "", 2, 1) = vbYes Then
                ProgressBar1.Value = 0
                Message_Label.Text = "Template01 생성중..."
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- Template Data Folder Name
                ls_ss = Split(STW_Template_Infomation(Form_Click_Index_Value, 4), "\")
                Template_Data_Folder_Name = ls_ss(UBound(ls_ss))

                Call Folder_Delete(sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name & "_" & date_value)

                fso.CopyFolder(sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4), sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\")
                Threading.Thread.Sleep(300)

                Dim Oil_Piping_folder = fso.GetFolder(sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name)
                Oil_Piping_folder.Name = Template_Data_Folder_Name & "_" & date_value
                Oil_Piping_folder_name = Oil_Piping_folder.Path
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- Folder 읽기 전용 변경
                Call wirte_transfer_file(Oil_Piping_folder_name)
                ProgressBar1.Value = 10
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
                Call Clash_Check_Folder_Create()
                ProgressBar1.Value = 20
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
                Call Data_Constraint_Delete(Oil_Piping_Initial_Constraint, Oil_Piping_Initial_Constraint_Count, Result_Template_Item_Value)
                '---------------------------------------------------------------------------------------------------
                ' 생 성 일  : -
                ' 생 성 자  : vb6에서 사용 된 함수
                ' 수 정 일  : 20.04.23
                ' 수 정 자  : 김원주
                ' 수정사유  : Call Oil_Piping_TCV_Constraint() 주석처리.
                '           : 터미네이트 에러 발생 예상 지점
                ' Parameter : -
                '---------------------------------------------------------------------------------------------------
                'Call Oil_Piping_TCV_Constraint()
                ProgressBar1.Value = 30
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
                Call Get_Parameter_Data_Comparison(Oil_Piping_Initial_Parameter, Oil_Piping_Initial_Parameter_Count, Me)
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- Oil_Piping Template Data Open
                Call CATIA_BASIC02()
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- Template Data Open
                Dim arrayOfVariantOfBSTR1(0)
                arrayOfVariantOfBSTR1(0) = Oil_Piping_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5)
                products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- Template Data Get Parameter
                '---------------------------------------------------------------------------------------------------
                ' 생 성 일  : -
                ' 생 성 자  : vb6에서 사용 된 함수
                ' 수 정 일  : 20.04.23
                ' 수 정 자  : 김원주
                ' 수정사유  : Call Oil_Piping_Filter_Get_Parameter() 주석처리.
                '           : SM2100기종에서 사용 안함.
                ' Parameter : -
                '---------------------------------------------------------------------------------------------------
                'Call Oil_Piping_Filter_Get_Parameter()
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- 기종변경
                '---------------------------------------------------------------------------------------------------
                ' 생 성 일  : -
                ' 생 성 자  : 
                ' 수 정 일  : 19.12.20
                ' 수 정 자  : 허혜원
                ' 수정사유  : 
                ' Parameter : -
                '---------------------------------------------------------------------------------------------------
                'selection.Clear
                'selection.search("Name='Oil_Piping_Machine_Type',all")
                'If Not selection.Count = 0 Then
                '    selection.Item(1).Value.Value = STW_Drawing_Infomation(1, 1)
                '    CATIA.ActiveDocument.Product.Update()
                'End If

                Dim Template_Skeleton_Part = Documents.Item(STW_Template_Infomation(Form_Click_Index_Value, 5)).Part
                'Call RuntimeCheck()
                '---------------------------------------------------------------------------------------------------- Input 변경
                For i = 1 To Input_Output_Excel_Value_Count - 1
                    'If "EPN06010205_P02" = Input_Output_Value(1, i) Then
                    '    MsgBox("", vbSystemModal)
                    'End If
                    If Input_Output_Value(4, i) = "D" Then
                        Call Constraint_Delete(Input_Output_Value(1, i), Input_Output_Value(2, i), Input_Output_Value(3, i))
                    ElseIf Input_Output_Value(4, i) = "A" Then
                        Call Input_Output_Axis_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), Template_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                    ElseIf Input_Output_Value(4, i) = "B" Then
                        Call Input_Output_Element_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), Template_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                        Template_Skeleton_Part.Update
                    ElseIf Input_Output_Value(4, i) = "C" Then
                        Call Input_Output_Initial_Parameter(Input_Output_Value(1, i), Input_Output_Value(2, i))
                        Template_Skeleton_Part.Update
                    ElseIf Input_Output_Value(4, i) = "I" Then
                        Call Input_Output_Excel_to_LayoutParameter(1, Input_Output_Value, i, 2, 1, O_BOM_Part_Value, 3, 1)
                        'Call RuntimeCheck()
                        '--------------------------------------------------------------------------------------------------- Oil_Tank Template TYPE
                    ElseIf Input_Output_Value(4, i) = "L" Then
                        Dim Layout_Parameter_Value = Me
                        Call Input_Output_UI_to_LayoutParameter_L(Layout_Parameter_Value, Input_Output_Value(1, i), Input_Output_Value(2, i))
                    ElseIf Input_Output_Value(4, i) = "G" Then
                        Call Input_Output_Insert_G_Type()
                    ElseIf Input_Output_Value(4, i) = "H" Then
                        'Call RuntimeCheck()
                        '--------------------------------------------------------------------------------------------------- Excel Open
                        Exc = CreateObject("excel.application")
                        Call Threading.Thread.Sleep(300)
                        'Exc.Visible = True
                        Exc.Workbooks.Open(Oil_Piping_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))

                        Call Input_Output_Excel_to_Excel(Exc, 1, Input_Output_Value(1, i), Input_Output_Value(2, i), O_BOM_Part_Value, 3, 1)

                        Exc.Application.DisplayAlerts = False
                        Exc.ActiveWorkbook.Save()
                        Exc.Quit()

                        Call KillProcess("EXCEL")
                    ElseIf Input_Output_Value(4, i) = "K" Then  '터미네이터??
                        'Call RuntimeCheck()
                        '--------------------------------------------------------------------------------------------------- Excel Open
                        Exc = CreateObject("excel.application")
                        Call Threading.Thread.Sleep(300)
                        'Exc.Visible = True
                        Exc.Workbooks.Open(Oil_Piping_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))

                        Call Input_Output_LayoutParameter_to_Excel(Exc, STW_Drawing_Infomation(1, 5), Input_Output_Value(1, i), Input_Output_Value(2, i))

                        Exc.Application.DisplayAlerts = False
                        Exc.ActiveWorkbook.Save()
                        Exc.Quit()

                        Call KillProcess("EXCEL")
                    End If
                    Threading.Thread.Sleep(200)
                    'Call RuntimeCheck()
                Next
                'Call RuntimeCheck()
                ProgressBar1.Value = 50
                '--------------------------------------------------------------------------------------------------- Bending Type(도면 Sheet삭제에 사용)
                '---------------------------------------------------------------------------------------------------
                ' 생 성 일  : -
                ' 생 성 자  : vb6에서 사용 된 함수
                ' 수 정 일  : 19.08.08
                ' 수 정 자  : 허혜원
                ' 수정사유  : 2012년 이후 미사용
                '           : 터미네이트 에러 발생 예상 지점
                ' Parameter : -
                '---------------------------------------------------------------------------------------------------
                'selection.search("Name='Bending_Type',all")
                'If Not selection.Count = 0 Then
                '    Bending_Type = selection.Item(1).Value
                'End If
                'Call RuntimeCheck()
                ''--------------------------------------------------------------------------------------------------- Oil_Piping_Hose_optimization
                '---------------------------------------------------------------------------------------------------
                ' 생 성 일  : -
                ' 생 성 자  : vb6에서 사용 된 함수
                ' 수 정 일  : 20.01.08
                ' 수 정 자  : 허혜원
                ' 수정사유  : SM2100에서 미사용
                '           : 터미네이트 에러 발생 예상 지점
                ' Parameter : -
                '---------------------------------------------------------------------------------------------------
                'If Strings.Left(assy_value_Name, 2) = "SM" And Not Strings.Right(assy_value_Name, 1) = "1" Then
                '    Call Oil_Piping_Hose_Optimization()
                'End If
                'Call RuntimeCheck()
                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- Oil_Piping Template Open
                Dim Oil_Piping_open = CATIA.Documents.Open(Oil_Piping_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5))
                ProgressBar1.Value = 70
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- VVIP Option Check
                If VVIP_Check.Checked = True Then
                    '--------------------------------------------------------------------------------------------------- Input 변경
                    For i = 1 To Input_Output_Excel_Value_Count - 1
                        If Input_Output_Value(4, i) = "M" Then
                            Call Input_Output_UI_to_LayoutParameter_M(Me, Input_Output_Value(1, i), Input_Output_Value(2, i))
                        End If
                    Next i

                    Oil_Piping_Parameter_Update.Enabled = True
                End If
                ProgressBar1.Value = 90
                'Call RuntimeCheck()
               Try
                    '--------------------------------------------------------------------------------------------------- PartNumber 변경
                    Threading.Thread.Sleep(200)
                    CATIA.DisplayFileAlerts = False
                    CATIA.ActiveDocument.Product.PartNumber = Template_Part_Number_Text.Text & "_" & date_value
                    Threading.Thread.Sleep(200)
                    'Call RuntimeCheck()
                    CATIA.ActiveDocument.Product.Update()
                    'Call RuntimeCheck()
                    CATIA.ActiveDocument.Save()
                    CATIA.DisplayFileAlerts = True
                Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                End Try
                'Call RuntimeCheck()
                '--------------------------------------------------------------------------------------------------- Oil Piping 간섭체크
                If Auto_Clash_Check.Checked = True Then
                    Call Oil_Piping_Clash_Check_Click(sender, e)
                    'Call RuntimeCheck()
                End If
                '--------------------------------------------------------------------------------------------------- Fit All In
                Threading.Thread.Sleep(500)
                CATIA.ActiveWindow.ActiveViewer.Reframe()
                Threading.Thread.Sleep(500)
                '--------------------------------------------------------------------------------------------------- Incident Report 창 종료
                ProgressBar1.Value = 100
                Message_Label.Text = "Template01 생성완료"
            End If
        Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Oil_Piping_Template_Create_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub Oil_Piping_Parameter_Control_Click(sender As Object, e As EventArgs) Handles Oil_Piping_Parameter_Control.Click
        '--------------------------------------------------------------------------------------------------- 화면 전환
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

        '--------------------------------------------------------------------------------------------------- Template_Parameter_Excel_Value 가져오기
        Call Get_Template_Parameter_Excel_Value(STW_COMP_Infomation(5, 6))

        '--------------------------------------------------------------------------------------------------- Parameter 창 Show
        ZB_Template_Parameter_Control.Show()

        Call Template_Parameter_Modify(STW_COMP_Infomation(5, 6))
        Call Folder_File_Count()
    End Sub


    Private Sub Oil_Piping_Clash_Check_Click(sender As Object, e As EventArgs) Handles Oil_Piping_Clash_Check.Click
        Message_Label.Text = "Oil Piping 간섭 체크중.."

        Call Template_Clash_Check_Function()

        Oil_Piping_Drawing_Create.Enabled = True
        Oil_Piping_Part_Convert.Enabled = True

        Message_Label.Text = "Oil Piping 간섭체크 완료"
    End Sub


    Private Sub Oil_Piping_Drawing_Create_Click(sender As Object, e As EventArgs) Handles Oil_Piping_Drawing_Create.Click
        Message_Label.Text = "Template02 생성중..."
        ProgressBar1.Value = 10
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change("Product", 1, 7)
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- Title Block 생성
        oil_piping_drawing_open = Documents.NewFrom(Oil_Piping_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 6))
        Call RuntimeCheck()
        Call CATIA_BASIC03()
        Call RuntimeCheck()
        DrawingDocument.Update()
        ProgressBar1.Value = 30
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- Layout Parameter -> Drawing
        For i = 1 To Input_Output_Excel_Value_Count - 1
            If Input_Output_Value(4, i) = "J" Then
                Call Input_Output_LayoutParameter_to_Drawing(Input_Output_Value(1, i), Input_Output_Value(2, i))
            End If
        Next i
        Call RuntimeCheck()
        '---------------------------------------------------------------------------------------------------
        Call Active_Window_Change("CATDrawing", 1, 10)
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 도면 정보에서 Item 가져오기
        Dim Drawing_Item_Folder_Value = Oil_Piping_folder_name
        Dim Drawing_Item_Count
        For Drawing_Item_Count = 1 To 20
            If STW_Template_Infomation(Form_Click_Index_Value, 1) = STW_Drawing_Infomation(0, Drawing_Item_Count) Then
                Exit For
            End If
        Next
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 도면 Part List 생성하기
        For i = 1 To STW_Drawing_Infomation_Count - 1
            If Not STW_Drawing_Infomation(i, Drawing_Item_Count) = "" Then
                DrawingSheets.Item(STW_Drawing_Infomation(i, Drawing_Item_Count)).Activate()
                Threading.Thread.Sleep(500)
                Call RuntimeCheck()
                Call Template_Title_Box_Create(Drawing_Item_Folder_Value & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7), STW_Drawing_Infomation(i, Drawing_Item_Count), i)
                Call RuntimeCheck()
            End If
        Next
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- Oil Piping 관련 특정 Parameter Setting
        If (Strings.Left(assy_value_Name, 2) = "SM3") Or (Strings.Left(assy_value_Name, 2) = "SM4") Or
            (Strings.Left(assy_value_Name, 2) = "SM5") Or (Strings.Left(assy_value_Name, 2) = "SM6") Then
            Call Oil_Piping_Drawing_Template_Setting()
        End If
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 도면 DWG_No 변경
        On Error Resume Next
        Drawingselection.Clear()
        Drawingselection.search("Name=Model*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = STW_Drawing_Infomation(1, 1)
        Next
        Drawingselection.Clear()
        '--------------------------------------------------------------------------------------------------- DWG_No
        Drawingselection.search("Name=DWG_No*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Template_Part_Number_Text.Text
        Next
        ProgressBar1.Value = 80
        On Error GoTo 0
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 표제란 입력
        Call drawing_title_box_write()
        DrawingViews.Item("Main View").Activate()
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 파일 Update
        Drawingselection.Clear()
        '--------------------------------------------------------------------------------------------------- SMx100
        If Strings.Mid(sLbl_ModelNumber, 4, 1) = "1" Then
            If Is_F_S = "S" Then
                'Drawingselection.Add(PIPING_MOTOR_OUTLET_D.e)
                'Drawingselection.Add(PIPING_MOTOR_OUTLET_N.D.e)
                'Drawingselection.Delete()
            End If
        Else
            If Is_F_S = "S" Then
                Drawingselection.Add(DrawingSheets.Item("HOSE MOTOR OIL RETURN DE"))
                Drawingselection.Add(DrawingSheets.Item("HOSE MOTOR OIL RETURN NDE"))

                If Drawingselection.Count > 0 Then
                    Drawingselection.Delete()
                End If
            End If
        End If
        Call RuntimeCheck()
        CATIA.ActiveDocument.Update()
        DrawingSheets.Item("PartList").Activate()
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- 도면 파일 저장
        CATIA.DisplayFileAlerts = False
        CATIA.ActiveDocument.SaveAs(Oil_Piping_folder_name & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATDrawing")
        CATIA.ActiveDocument.Save()
        CATIA.DisplayFileAlerts = True
        Call RuntimeCheck()
        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()
        Call RuntimeCheck()
        ProgressBar1.Value = 100
        Message_Label.Text = "Template02 생성완료"
    End Sub


    Private Sub Oil_Piping_Part_Convert_Click(sender As Object, e As EventArgs) Handles Oil_Piping_Part_Convert.Click
        Try
            ProgressBar1.Value = 10
            Message_Label.Text = "Template03 생성중..."

            '--------------------------------------------------------------------------------------------------- 화면 변경
            Call CATIA_BASIC02()
            Call Active_Window_Change("CATPart", 1, 7)
            ProgressBar1.Value = 30

            '--------------------------------------------------------------------------------------------------- Product File Save
            CATIA.DisplayFileAlerts = False
            CATIA.ActiveDocument.Product.Update()
            CATIA.ActiveDocument.Save()

            '--------------------------------------------------------------------------------------------------- 동일한 이름의 Data가 있는지 확인
            Call Save_File_Check(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")

            '--------------------------------------------------------------------------------------------------- Part File Save
            CATIA.ActiveDocument.SaveAs(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            CATIA.ActiveDocument.Close()
            CATIA.DisplayFileAlerts = True

            '--------------------------------------------------------------------------------------------------- 도면 파일 Close
            Call CATIA_BASIC02()
            Call Active_Window_Change("CATDrawing", 1, 10)
            If Strings.Right(CATIA.ActiveDocument.Name, 10) = "CATDrawing" Then
                CATIA.ActiveDocument.Close()
            End If

            '--------------------------------------------------------------------------------------------------- Oil_Piping EPN 변경
            Call CATIA_BASIC02()
            If Strings.Left(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name, 1) <> "E" Then
                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value
            End If

            '--------------------------------------------------------------------------------------------------- Product 저장
            Dim oil_piping_product = Documents.Item(Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            oil_piping_product.Save()

            '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            Call O_BOM_Excel_Writing(Result_Template_Item_Value, Template_Part_Number_Text.Text & "_" & date_value)
            Call Btn_Refresh_Click()

            '--------------------------------------------------------------------------------------------------- 파일 속성 변경 (읽기 전용)
            fso = CreateObject("Scripting.FileSystemObject")
            Dim file_Attributes_value = fso.GetFile(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            file_Attributes_value.Attributes = 33

            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            ProgressBar1.Value = 100
            Message_Label.Text = "Template03 생성완료"
        Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Oil_Piping_Part_Convert_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub FILTER_MOVE_H_LENGTH_Update_Click(sender As Object, e As EventArgs) Handles FILTER_MOVE_H_LENGTH_Update.Click
        If filter_move_h_length = Nothing Then
            Call Oil_Piping_Filter_Get_Parameter()
        End If

        If Not filter_move_h_length = Nothing Then
            If filter_move_h_length_text.Text = "" Then
                Call SHOW_MESSAGE_BOX("입력값이 없습니다.", "", 1, 1)
                Exit Sub
            End If

            filter_move_h_length.Value = filter_move_h_length_text.Text
            CATIA.ActiveDocument.Product.Update()

            Threading.Thread.Sleep(300)
        End If
    End Sub


    Private Sub FILTER_SIDE_DIRECTION_MOVE_Update_Click(sender As Object, e As EventArgs) Handles FILTER_SIDE_DIRECTION_MOVE_Update.Click
        If filter_move_h_length = Nothing Then
            Call Oil_Piping_Filter_Get_Parameter()
        End If

        If Not filter_side_direction_move = Nothing Then
            If filter_side_direction_move_text.Text = "" Then
                Call SHOW_MESSAGE_BOX("입력값이 없습니다.", "", 1, 1)
                Exit Sub
            End If

            filter_side_direction_move.Value = filter_side_direction_move_text.Text
            CATIA.ActiveDocument.Product.Update()

            Threading.Thread.Sleep(300)
        End If
    End Sub


    Private Sub FILTER_MOVE_V_LENGTH_Update_Click(sender As Object, e As EventArgs) Handles FILTER_MOVE_V_LENGTH_Update.Click
        If filter_move_h_length = Nothing Then
            Call Oil_Piping_Filter_Get_Parameter()
        End If

        If Not filter_move_v_length = Nothing Then
            If filter_move_v_length_text.Text = "" Then
                Call SHOW_MESSAGE_BOX("입력값이 없습니다.", "", 1, 1)
                Exit Sub
            End If

            filter_move_v_length.Value = filter_move_v_length_text.Text
            CATIA.ActiveDocument.Product.Update()

            Threading.Thread.Sleep(300)
        End If
    End Sub


    Private Sub Oil_Piping_Parameter_Update_Click(sender As Object, e As EventArgs) Handles Oil_Piping_Parameter_Update.Click
        Message_Label.Text = "Update 진행중"

        '--------------------------------------------------------------------------------------------------- Input 변경
        For i = 1 To Input_Output_Excel_Value_Count - 1
            If Input_Output_Value(4, i) = "M" Then
                Call Input_Output_UI_to_LayoutParameter_M(Me, Input_Output_Value(1, i), Input_Output_Value(2, i))
            End If
        Next i

        '--------------------------------------------------------------------------------------------------- Product File Save
        CATIA.ActiveDocument.Product.Update()
        CATIA.ActiveDocument.Save()

        Message_Label.Text = "Update 완료"
    End Sub


    Private Sub VVIP_Check_CheckedChanged(sender As Object, e As EventArgs) Handles VVIP_Check.CheckedChanged
        Call VVIP_Check_Form_Size(362, 532, Me)
    End Sub

    Private Sub OilPipe_Delay_Text_KeyPress(sender As Object, e As KeyPressEventArgs) Handles OilPipe_Delay_Text.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
End Class