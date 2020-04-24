Public Class CA_Coupling_Cover_Template_Create
    Public LENG_COVER_HOLE_SPACING, DIA_COVER_HOLE, DIA_COVER_HOLE_MAX


    Private Sub CA_Coupling_Cover_Template_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- Form Position
        Me.Top = 150
        Me.Left = 50
        Me.Width = 311
        Me.Height = 474
        Me.TopMost = True

        Call CATIA_Basic02()

        UI_COVER_EXT.Enabled = True
        UI_STAGE_COMBO.Enabled = False
        Coupling_Cover_Drawing_Create.Enabled = False
        Coupling_Cover_Result_Part_Create.Enabled = False

        '--------------------------------------------------------------------------------------------------- BOM Path 및 날자 가져오기
        Call Data_O_BOM_Excel_Search()

        '--------------------------------------------------------------------------------------------------- Coupling_Cover_Excel_Open
        Call Coupling_Cover_Excel_Open()

        '--------------------------------------------------------------------------------------------------- Coupling_Cover_Template_Form_Inicial_Setting
        Call Coupling_Cover_Template_Form_Inicial_Setting()

        '--------------------------------------------------------------------------------------------------- 도면 관련 정보 가져오기
        Call Get_Excel_Value_Function(Application.StartupPath & "\STW_Drawing_Infomation.xlsx", assy_value_Name, 9, _
                                      "A", "B", "C", "D", "E", "F", "G", "H", "I", "", STW_Drawing_Infomation_Count, STW_Drawing_Infomation, "5")

        '--------------------------------------------------------------------------------------------------- Form Position
        Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)

        '--------------------------------------------------------------------------------------------------- SMx100
        If Mid(sLbl_ModelNumber, 4, 1) = "1" Then
            UI_COVER_EXT.Enabled = False
        End If

        '--------------------------------------------------------------------------------------------------- 단수 활성화
        If Strings.Left(sLbl_ModelNumber, 2) = "SM" And Not Strings.Mid(sLbl_ModelNumber, 4, 1) = "1" Then
            UI_STAGE_COMBO.Enabled = True
        ElseIf Strings.Left(sLbl_ModelNumber, 3) = "SME" Then
            UI_STAGE_COMBO.Enabled = True
        End If

        '--------------------------------------------------------------------------------------------------- Stage_Combo 초기값 입력
        If UI_STAGE_COMBO.Enabled = True Then
            UI_STAGE_COMBO.Text = UI_STAGE_COMBO.Items.Item(0)
        Else
            UI_STAGE_COMBO.Text = ""
        End If
    End Sub


    Private Sub VVIP_Check_CheckedChanged(sender As Object, e As EventArgs) Handles VVIP_Check.CheckedChanged
        Call VVIP_Check_Form_Size(311, 541, Me)
    End Sub


    Private Sub Coupling_Cover_Parameter_Control_Click(sender As Object, e As EventArgs) Handles Coupling_Cover_Parameter_Control.Click
        '--------------------------------------------------------------------------------------------------- 화면 전환
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

        '--------------------------------------------------------------------------------------------------- Template_Parameter_Excel_Value 가져오기
        Call Get_Template_Parameter_Excel_Value(STW_COMP_Infomation(3, 6))

        '--------------------------------------------------------------------------------------------------- Parameter 창 Show
        ZB_Template_Parameter_Control.Show()

        Call Template_Parameter_Modify(STW_COMP_Infomation(3, 6))
        Call Folder_File_Count()
    End Sub


    Private Sub Coupling_Cover_Template_Create_Click(sender As Object, e As EventArgs) Handles Coupling_Cover_Template_Create.Click
        Try
            ProgressBar1.Value = 0

            '--------------------------------------------------------------------------------------------------- Coupling Cover Part No. 가 없을때
            If Template_Part_Number_Text.Text = "" Then
                Call SHOW_MESSAGE_BOX("Coupling Cover Part No.가 없습니다.", "", 1, 1)
                Exit Sub
            End If
            '--------------------------------------------------------------------------------------------------- Template 실행여부 확인
            If SHOW_MESSAGE_BOX("Template01 실행 하겠습니까?", "", 2, 1) = vbYes Then
                Message_Label.Text = "Template01 생성중..."

                '--------------------------------------------------------------------------------------------------- Template Data Folder Name
                ls_ss = Split(STW_Template_Infomation(Form_Click_Index_Value, 4), "\")
                Template_Data_Folder_Name = ls_ss(UBound(ls_ss))

                Call Folder_Delete(sProject_DB_Path_text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name & "_" & date_value)

                fso.CopyFolder(sTemplate_Data_Path_text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4), _
                               sProject_DB_Path_text & "\" & Sales_No & "_" & assy_value_Name & "\")

                Threading.Thread.Sleep(500)

                Dim coupling_cover_folder = fso.getfolder(sProject_DB_Path_text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name)
                coupling_cover_folder.Name = Template_Data_Folder_Name & "_" & date_value
                coupling_cover_folder_name = coupling_cover_folder.Path

                '--------------------------------------------------------------------------------------------------- Folder 읽기 전용 변경
                Call wirte_transfer_file(coupling_cover_folder_name)
                ProgressBar1.Value = 10

                '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
                Call Clash_Check_Folder_Create()
                ProgressBar1.Value = 20

                '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
                Call Data_Constraint_Delete(Coupling_Cover_Initial_Constraint, Coupling_Cover_Initial_Constraint_Count, Result_Template_Item_Value)
                ProgressBar1.Value = 30

                '--------------------------------------------------------------------------------------------------- Coupling Cover Template Data Open
                Call CATIA_BASIC02()

                '--------------------------------------------------------------------------------------------------- Template Data Open
                Dim arrayOfVariantOfBSTR1(0)
                arrayOfVariantOfBSTR1(0) = coupling_cover_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5)
                products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

                Dim Template_Skeleton_Part = Nothing
                Template_Skeleton_Part = Documents.Item("COUPLING_COVER_Template_Skeleton.CATPart").Part

                If Template_Skeleton_Part Is Nothing Then
                    Template_Skeleton_Part = Documents.Item("SM21_COUPLING_COVER_TEMPLATE_SKELETON.CATPart").Part
                End If

                '--------------------------------------------------------------------------------------------------- Input 변경
                'A:              Axis Replace
                'B:              Element Replace
                'C:              Layout Parameter -> Template Parameter
                'D:              Position Constraint
                'E:              Output Axis
                'F:              Output Parameter
                'G:              커플링BOM -> Parameter
                'H:              BOM P / N -> Excel
                'I:              BOM P / N -> Parameter
                'J:              Layout Parameter -> 도면Template
                'K:              Layout Parameter -> 도면엑셀
                'L:              UI에 있는 값 -> Template Parameter
                'M:              UI에 VVIP 적용시 -> Template Parameter
                For i = 1 To Input_Output_Excel_Value_Count - 1
                    Debug.Print(Input_Output_Value(1, i))

                    If Input_Output_Value(4, i) = "D" Then
                        Call Constraint_Delete(Input_Output_Value(1, i), Input_Output_Value(2, i), Input_Output_Value(3, i))
                    ElseIf Input_Output_Value(4, i) = "A" Then
                        Call Input_Output_Axis_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), Template_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                    ElseIf Input_Output_Value(4, i) = "B" Then
                        Call Input_Output_Element_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), Template_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                    ElseIf Input_Output_Value(4, i) = "C" Then
                        Call Input_Output_Initial_Parameter(Input_Output_Value(1, i), Input_Output_Value(2, i))
                    ElseIf Input_Output_Value(4, i) = "I" Then
                        Call Input_Output_Excel_to_LayoutParameter(1, Input_Output_Value, i, 2, 1, O_BOM_Part_Value, 3, 1)
                    ElseIf Input_Output_Value(4, i) = "L" Then
                        Dim Layout_Parameter_Value = Me
                        Call Input_Output_UI_to_LayoutParameter_L(Layout_Parameter_Value, Input_Output_Value(1, i), Input_Output_Value(2, i))
                    ElseIf Input_Output_Value(4, i) = "G" Then
                        Call Input_Output_Insert_G_Type()
                    ElseIf Input_Output_Value(4, i) = "H" Then
                        '--------------------------------------------------------------------------------------------------- Excel Open
                        If Exc.Name = "" Then
                            Exc = CreateObject("excel.application")
                            Exc.Workbooks.Open(coupling_cover_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))
                            'Exc.Visible = True
                        End If

                        Call Input_Output_Excel_to_Excel(Exc, 1, Input_Output_Value(1, i), Input_Output_Value(2, i), O_BOM_Part_Value, 3, 1)

                        If Not Input_Output_Value(4, i + 1) = "H" Then
                            Exc.AlertBeforeOverwriting = False      ' 저장시 덮어쓰기
                            Exc.ActiveWorkbook.Save()
                            Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()
                            Exc.Quit()
                        End If
                    ElseIf Input_Output_Value(4, i) = "K" Then
                        '--------------------------------------------------------------------------------------------------- Excel Open
                        If Exc.Name = "" Then
                            Exc = CreateObject("excel.application")
                            Exc.Workbooks.Open(coupling_cover_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7))
                            'Exc.Visible = True
                        End If

                        Call Input_Output_LayoutParameter_to_Excel(Exc, "PartList", Input_Output_Value(1, i), Input_Output_Value(2, i))

                        If Not Input_Output_Value(4, i + 1) = "K" Then
                            Exc.AlertBeforeOverwriting = False      ' 저장시 덮어쓰기
                            Exc.ActiveWorkbook.Save()
                            Exc.Application.DisplayAlerts = False : Exc.ActiveWorkbook.Close()
                            Exc.Quit()
                        End If
                    End If
                Next

                ProgressBar1.Value = 50
                selection.Clear()
                '--------------------------------------------------------------------------------------------------- Machine_Type Parameter 변경
                selection.search("Name='Machine_Type',all")
                selection.Item(1).Value.Value = STW_Drawing_Infomation(1, 1)
                Template_Skeleton_Part.Update()
                CATIA.ActiveDocument.Product.Update()
                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value       'hwheo 180718 수정

                '--------------------------------------------------------------------------------------------------- Coupling Cover Template Open
                Dim coupling_cover_open = CATIA.Documents.Open(coupling_cover_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5))

                '--------------------------------------------------------------------------------------------------- Template Data Get Parameter
                Call Template_Data_Get_Parameter()

                ProgressBar1.Value = 70

                '--------------------------------------------------------------------------------------------------- VVIP Option Check
                If VVIP_Check.Checked = True Then
                    '--------------------------------------------------------------------------------------------------- Input 변경
                    For i = 1 To Input_Output_Excel_Value_Count - 1
                        If Input_Output_Value(4, i) = "M" Then
                            Call Input_Output_UI_to_LayoutParameter_M(Me, Input_Output_Value(1, i), Input_Output_Value(2, i))
                        End If
                    Next i
                    CATIA.ActiveDocument.Product.Update()
                    Threading.Thread.Sleep(1500)
                    selection.Clear()
                    selection.search("Name='LENG_COVER_HOLE_SPACING',all")
                    If Not selection.Count = 0 Then
                        LENG_COVER_HOLE_SPACING = selection.Item(1).Value
                        LENG_COVER_HOLE_SPACING_Text.Text = LENG_COVER_HOLE_SPACING.Value
                    End If
                    selection.Clear()
                    selection.search("Name='DIA_COVER_HOLE',all")
                    If Not selection.Count = 0 Then
                        DIA_COVER_HOLE = selection.Item(1).Value
                        DIA_COVER_HOLE_Text.Text = DIA_COVER_HOLE.Value
                    End If
                    selection.Clear()
                    selection.search("Name='DIA_COVER_HOLE_MAX',all")
                    If Not selection.Count = 0 Then
                        DIA_COVER_HOLE_MAX = selection.Item(1).Value
                        HOLE_DIA_MAX_Label.Text = DIA_COVER_HOLE_MAX.Value
                    End If

                    Coupling_Cover_Parameter_Update.Enabled = True
                End If

                ProgressBar1.Value = 90
                '--------------------------------------------------------------------------------------------------- Product File Save
                CATIA.ActiveDocument.Product.Update()
                Threading.Thread.Sleep(1000)

                CATIA.ActiveDocument.Save()
            Else
                Exit Sub
            End If

            CATIA.ActiveDocument.Product.Update()
            CATIA.ActiveDocument.Product.Update()
            '--------------------------------------------------------------------------------------------------- Coupling Cover 간섭체크
            If Auto_Clash_Check.Checked = True Then
                Call Coupling_Cover_Clash_Check_Click(Me, Nothing)
            End If
            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            ProgressBar1.Value = 100
            Message_Label.Text = "Template01 생성완료"
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            ProgressBar1.Value = 100
            Message_Label.Text = "Template01 생성 오류"
        End Try
    End Sub


    Private Sub Coupling_Cover_Clash_Check_Click(sender As Object, e As EventArgs) Handles Coupling_Cover_Clash_Check.Click
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Message_Label.Text = "Coupling Cover 간섭 체크중.."
        Call Template_Clash_Check_Function()

        Coupling_Cover_Drawing_Create.Enabled = True
        Coupling_Cover_Result_Part_Create.Enabled = True

        Message_Label.Text = "Coupling Cover 간섭체크 완료"
    End Sub


    Private Sub Coupling_Cover_Drawing_Create_Click(sender As Object, e As EventArgs) Handles Coupling_Cover_Drawing_Create.Click

        ProgressBar1.Value = 0
        Message_Label.Text = "Template02 생성중..."
        ProgressBar1.Value = 10

        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

        '--------------------------------------------------------------------------------------------------- COMP 저장
        CATIA.ActiveDocument.Save()

        '--------------------------------------------------------------------------------------------------- Title Block 생성
        Dim coupling_cover_drawing_open = CATIA.Documents.NewFrom(coupling_cover_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 6))
        ProgressBar1.Value = 20

        Call CATIA_BASIC03()
        ProgressBar1.Value = 30

        '--------------------------------------------------------------------------------------------------- 도면 정보에서 Item 가져오기
        Dim Drawing_Item_Folder_Value = coupling_cover_folder_name

        Dim Drawing_Item_Count
        For Drawing_Item_Count = 1 To 20
            If STW_Template_Infomation(Form_Click_Index_Value, 1) = STW_Drawing_Infomation(0, Drawing_Item_Count) Then
                Exit For
            End If
        Next

        '--------------------------------------------------------------------------------------------------- 도면 Part List 생성하기
        For i = 1 To STW_Drawing_Infomation_Count - 1
            If Not STW_Drawing_Infomation(i, Drawing_Item_Count) = "" Then
                DrawingSheets.Item(STW_Drawing_Infomation(i, Drawing_Item_Count)).Activate()
                Call Template_Title_Box_Create(Drawing_Item_Folder_Value & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7), STW_Drawing_Infomation(i, Drawing_Item_Count), i)
            End If
        Next
        ProgressBar1.Value = 60

        On Error Resume Next
        Drawingselection.Clear()
        '--------------------------------------------------------------------------------------------------- 도면 DWG_No 변경
        Drawingselection.search("Name=Model*,all")
        For j = 1 To Drawingselection.Count
            Drawingselection.Item(j).Value.Text = STW_Drawing_Infomation(1, 1)
        Next
        Drawingselection.Clear()
        '--------------------------------------------------------------------------------------------------- DWG_No
        Drawingselection.search("Name=DWG_No*,all")
        For j = 1 To Drawingselection.Count
            Drawingselection.Item(j).Value.Text = Template_Part_Number_Text.Text
        Next
        On Error GoTo 0

        ProgressBar1.Value = 80

        '--------------------------------------------------------------------------------------------------- 표제란 입력
        Call drawing_title_box_write()

        '--------------------------------------------------------------------------------------------------- Layout Parameter -> Drawing
        For i = 1 To Input_Output_Excel_Value_Count - 1
            If Input_Output_Value(4, i) = "J" Then
                Call Input_Output_LayoutParameter_to_Drawing(Input_Output_Value(1, i), Input_Output_Value(2, i))
            End If
        Next i

        '--------------------------------------------------------------------------------------------------- 파일 Update
        CATIA.ActiveDocument.Update()

        '--------------------------------------------------------------------------------------------------- 도면 파일 저장
        CATIA.ActiveDocument.SaveAs(coupling_cover_folder_name & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATDrawing")
        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        ProgressBar1.Value = 100
        Message_Label.Text = "Template02 생성완료"
    End Sub


    Private Sub Coupling_Cover_Result_Part_Create_Click(sender As Object, e As EventArgs) Handles Coupling_Cover_Result_Part_Create.Click
        Try
            Message_Label.Text = "Template03 생성중..."
            ProgressBar1.Value = 0

            '--------------------------------------------------------------------------------------------------- 화면 변경
            Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

            '--------------------------------------------------------------------------------------------------- Product -> Part 변환
            CATIA.ActiveDocument.Save()
            Call Generate_CATPart_from_Product()

            '--------------------------------------------------------------------------------------------------- Geometrycal 이름에 "\" 기호 삭제
            Call Delete_Symbol()

            ProgressBar1.Value = 10

            '--------------------------------------------------------------------------------------------------- PartNumber 변경
            CATIA.ActiveDocument.Product.PartNumber = Template_Part_Number_Text.Text & "_" & date_value





            Dim s1
            s1 = CATIA.ActiveDocument.Product.publications
            For i = 1 To Input_Output_Excel_Value_Count - 1
                '--------------------------------------------------------------------------------------------------- EPN04_P01 Publication 생성
                If Input_Output_Value(4, i) = "E" Then
                    Call publication_axis_create(Input_Output_Value(1, i))
                    '--------------------------------------------------------------------------------------------------- EPN04_D01 parameter 생성
                ElseIf Input_Output_Value(4, i) = "F" Then
                    Call part_parameter_create(Input_Output_Value(1, i), Input_Output_Parameter_Value(i), Input_Output_Value(2, i))
                End If
            Next

            '--------------------------------------------------------------------------------------------------- 동일한 이름의 Data가 있는지 확인
            Call Save_File_Check(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & _
                                 Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            CATIA.DisplayFileAlerts = False
            '--------------------------------------------------------------------------------------------------- Part File Save
            CATIA.ActiveDocument.SaveAs(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & _
                                 Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            CATIA.DisplayFileAlerts = True
            '--------------------------------------------------------------------------------------------------- Part Data Close
            CATIA.ActiveDocument.Close()

            '--------------------------------------------------------------------------------------------------- 도면 파일 Close
            Call CATIA_Basic02()
            Call Active_Window_Change("CATDrawing", 1, 10)
            If Strings.Right(CATIA.ActiveDocument.Name, 10) = "CATDrawing" Then
                CATIA.ActiveDocument.Close()
            End If

            '--------------------------------------------------------------------------------------------------- Template 파일 Close
            Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
            If CATIA.ActiveDocument.Name = STW_Template_Infomation(Form_Click_Index_Value, 5) Then
                CATIA.ActiveDocument.Close()
            End If
            ProgressBar1.Value = 30

            '--------------------------------------------------------------------------------------------------- COMP 에서 Template Data 제거
            Call CATIA_Basic02()
            selection.Clear()
            selection.Add(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count))
            selection.Delete()

            '--------------------------------------------------------------------------------------------------- COMP Data에 실적 DATA 삽입
            Threading.Thread.Sleep(500)
            Dim arrayOfVariantOfBSTR1(0)
            arrayOfVariantOfBSTR1(0) = sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & _
                                       Template_Part_Number_Text.Text & "_" & date_value & ".CATPart"
            CATIA.ActiveDocument.Product.products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            '--------------------------------------------------------------------------------------------------- Coupling Cover Part Name 변경
            If Strings.Left(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name, 1) <> "E" Then
                CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value
            End If

            '--------------------------------------------------------------------------------------------------- Coupling Cove Data 저장
            Dim coupling_cover_part = Documents.Item(Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            coupling_cover_part.Save()
            ProgressBar1.Value = 60

            '--------------------------------------------------------------------------------------------------- coupling_cover Constraint Delete & Create
            For i = 1 To Coupling_Cover_Selection_Constraint_Count - 1
                Call Constraint_Delete(Coupling_Cover_Selection_Constraint(i, 1), Coupling_Cover_Selection_Constraint(i, 2), Coupling_Cover_Selection_Constraint(i, 3))
            Next

            '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
            Call O_BOM_Excel_Writing(Result_Template_Item_Value, Template_Part_Number_Text.Text & "_" & date_value)
            Call Btn_Refresh_Click()

            '--------------------------------------------------------------------------------------------------- 파일 속성 변경 (읽기 전용)
            Threading.Thread.Sleep(500)
            fso = CreateObject("Scripting.FileSystemObject")
            Dim file_Attributes_value = fso.GetFile(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & _
                                                Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
            file_Attributes_value.Attributes = 33

            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            ProgressBar1.Value = 100
            Message_Label.Text = "Template03 생성완료"
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            ProgressBar1.Value = 100
            Message_Label.Text = "Template03 생성 오류"
        End Try
    End Sub


    Private Sub Coupling_Cover_Parameter_Update_Click(sender As Object, e As EventArgs) Handles Coupling_Cover_Parameter_Update.Click
        Try
            Message_Label.Text = "Update 진행 중..."

            Call CATIA_BASIC02()
            selection.Clear()
            selection.search("Name='LENG_COVER_HOLE_SPACING',all")
            If Not selection.Count = 0 Then
                LENG_COVER_HOLE_SPACING = selection.Item(1).Value
            End If
            selection.Clear()
            selection.search("Name='DIA_COVER_HOLE',all")
            If Not selection.Count = 0 Then
                DIA_COVER_HOLE = selection.Item(1).Value
            End If

            selection.Clear()

            Call VVIP_Option(Me.VVIP_Check, Me.UI_COVER_EXT, "EPN04_VVIP_OPTION")
            Call VVIP_Option(Me.VVIP_Check, Me.UI_PUNCH_HOLE, "타공생성")

            '--------------------------------------------------------------------------------------------------- Hole Parameter 변경
            Call Parameter_Error_Message(1, DIA_COVER_HOLE_MAX.Value, DIA_COVER_HOLE_Text, DIA_COVER_HOLE)

            If Not LENG_COVER_HOLE_SPACING Is Nothing Then
                LENG_COVER_HOLE_SPACING.Value = LENG_COVER_HOLE_SPACING_Text.Text
            End If

            If Not DIA_COVER_HOLE Is Nothing Then
                DIA_COVER_HOLE.Value = DIA_COVER_HOLE_Text.Text
            End If

            '--------------------------------------------------------------------------------------------------- Product File Save
            CATIA.ActiveDocument.Product.Update()
            CATIA.ActiveDocument.Save()

            Message_Label.Text = "Update 완료"
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Coupling_Cover_Parameter_Update_Click() 함수 오류!!", "", 1, 1)
        End Try
    End Sub

End Class