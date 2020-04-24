Public Class FA_AC_Outlet_Template_Create

    Private Sub FA_AC_Outlet_Template_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '--------------------------------------------------------------------------------------------------- 초기 Form 크기 적용 ( VVIP Optioin )
        Me.Top = 150
        Me.Left = 50
        Me.Width = 473 ' 652
        Me.Height = 696
        Me.TopMost = True

        AC_Outlet_Drawing_Create.Enabled = False
        AC_Outlet_Result_Part_Create.Enabled = False

        Call CATIA_BASIC02()

        '--------------------------------------------------------------------------------------------------- BOM Path 및 날자 가져오기
        Call Data_O_BOM_Excel_Search()

        '--------------------------------------------------------------------------------------------------- AC_Outlet_Excel_Open
        Call AC_Outlet_Excel_Open()

        '--------------------------------------------------------------------------------------------------- 도면 관련 정보 가져오기
        sExcelFilePath = Application.StartupPath & "\STW_Drawing_Infomation.xlsx"
        If System.IO.File.Exists(sExcelFilePath) Then
            Call Get_Excel_Value_Function(sExcelFilePath, assy_value_Name, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "", _
                                          STW_Drawing_Infomation_Count, STW_Drawing_Infomation, "5")
        Else
            Call SHOW_MESSAGE_BOX("엑셀 파일이 없어 종료합니다." & vbCrLf & "-> " & sExcelFilePath, "", 1, 1)
            Exit Sub
        End If

        Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)

        '--------------------------------------------------------------------------------------------------- Input_Output Value  가져오기
        Call Input_Output_Excel_Value(STW_Template_Infomation(Form_Click_Index_Value, 1))

        '--------------------------------------------------------------------------------------------------- Scroll 초기화
        VScroll1.Maximum = -10
        VScroll1.Minimum = 10
        VScroll1.LargeChange = 1
        VScroll1.SmallChange = 1

        VScroll2.Maximum = -10
        VScroll2.Minimum = 10
        VScroll2.LargeChange = 1
        VScroll2.SmallChange = 1

        '--------------------------------------------------------------------------------------------------- Stage_Combo 초기화
        Stage_Combo.Items.Clear()
        Stage_Combo.Items.Add("2")
        Stage_Combo.Items.Add("3")

        Type_Combo.Items.Clear()
        Type_Combo.Items.Add("1")
    End Sub


    Private Sub AC_Outlet_Template_Create_Click(sender As Object, e As EventArgs) Handles AC_Outlet_Template_Create.Click
        '--------------------------------------------------------------------------------------------------- AC Outlet Part NO 가 없을때
        If Template_Part_Number_Text.Text = "" Then
            Call SHOW_MESSAGE_BOX("AC Outlet Part No.가 없습니다.", "", 1, 1)
            Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- Template 실행여부 확인
        If SHOW_MESSAGE_BOX("Template01 실행 하겠습니까?", "", 2, 1) = vbYes Then
            ProgressBar1.Value = 0
            Message_Label.Text = "Template01 생성중..."

            '--------------------------------------------------------------------------------------------------- Template Data Folder Name
            ls_ss = Split(STW_Template_Infomation(Form_Click_Index_Value, 4), "\")
            Template_Data_Folder_Name = ls_ss(UBound(ls_ss))
            Call Folder_Delete(sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name & "_" & date_value)
            fso.CopyFolder(sTemplate_Data_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 4), sProject_DB_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\")

            Dim AC_Outlet_folder = fso.getfolder(sTemplate_Data_Path_Text & "\" & Sales_No & "_" & assy_value_Name & "\" & Template_Data_Folder_Name)
            AC_Outlet_folder.Name = Template_Data_Folder_Name & "_" & date_value
            AC_Outlet_folder_name = AC_Outlet_folder

            '--------------------------------------------------------------------------------------------------- Folder 읽기 전용 변경
            Call wirte_transfer_file(AC_Outlet_folder_name)
            ProgressBar1.Value = 10

            '--------------------------------------------------------------------------------------------------- 간섭 체크 생성 폴더 확인
            Call Clash_Check_Folder_Create()
            ProgressBar1.Value = 20

            '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
            Call Data_Constraint_Delete(AC_Outlet_Initial_Constraint, AC_Outlet_Initial_Constraint_Count, Result_Template_Item_Value)
            ProgressBar1.Value = 30

            '--------------------------------------------------------------------------------------------------- Template에서 실적파트와 비교 대상 측정
            Call Get_Parameter_Data_Comparison(AC_Outlet_Initial_Parameter, AC_Outlet_Initial_Parameter_Count, Me)

            '--------------------------------------------------------------------------------------------------- AC_Outlet Template Data Open
            Call CATIA_BASIC02()
            CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value       'hwheo 180718 수정

            '--------------------------------------------------------------------------------------------------- Template Data Open
            Dim arrayOfVariantOfBSTR1(0)
            arrayOfVariantOfBSTR1(0) = AC_Outlet_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 5)
            products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            Dim AC_Outlet_Skeleton_Part = Documents.Item("AC_Outlet_Pipe_Skeleton.CATPart").Part

            '--------------------------------------------------------------------------------------------------- Input 변경
            For i = 1 To Input_Output_Excel_Value_Count - 1
                If Input_Output_Value(4, i) = "D" Then
                    Call Constraint_Delete(Input_Output_Value(1, i), Input_Output_Value(2, i), Input_Output_Value(3, i))
                ElseIf Input_Output_Value(4, i) = "A" Then
                    Call Input_Output_Axis_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), AC_Outlet_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                ElseIf Input_Output_Value(4, i) = "B" Then
                    Call Input_Output_Element_Replace(Input_Output_Value(1, i), Input_Output_Value(2, i), AC_Outlet_Skeleton_Part, "WF_INPUT", "WF_TR_INPUT")
                ElseIf Input_Output_Value(4, i) = "C" Then
                    Call Input_Output_Initial_Parameter(Input_Output_Value(1, i), Input_Output_Value(2, i))
                End If
            Next
            ProgressBar1.Value = 50

            '--------------------------------------------------------------------------------------------------- BOM_Excel 정보 가져오기(PYO?)
            Call Get_Excel_Value_Function(A_Main_Form.re_o_bom_path.Text & "\" & A_Main_Form.O_BOM_File_Name.Text,
                                          "BOM_LIST", 4, "B", "H", "K", "J", "", "", "", "", "", "", O_BOM_Part_Value_Count, O_BOM_Part_Value, "1")

            Call Ac_Outlet_Data_Replace_Excel()
            ProgressBar1.Value = 70

            '--------------------------------------------------------------------------------------------------- Template Data Replace
            Call Ac_Outlet_Data_Replace()
            ProgressBar1.Value = 80

            '--------------------------------------------------------------------------------------------------- Template Data Get Parameter
            Call Template_Data_Get_Parameter()
            ProgressBar1.Value = 90

            '--------------------------------------------------------------------------------------------------- Product File Save
            CATIA.ActiveDocument.Product.Update()
            CATIA.ActiveDocument.Save()
        Else
            Exit Sub
        End If

        CATIA.ActiveDocument.Product.Update()
        CATIA.ActiveDocument.Product.Update()

        '--------------------------------------------------------------------------------------------------- AC Outlet 간섭체크
        If Auto_Clash_Check.Checked = True Then
            Call AC_Outlet_Clash_Check_Click(sender, Nothing)
        End If

        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        ProgressBar1.Value = 100
        Message_Label.Text = "Template01 생성완료"
    End Sub


    Private Sub AC_Outlet_Clash_Check_Click(sender As Object, e As EventArgs) Handles AC_Outlet_Clash_Check.Click
        '--------------------------------------------------------------------------------------------------- Data 창 전환 ( COMP Assy )
        Message_Label.Text = "AC Outlet 간섭 체크중.."

        Call Template_Clash_Check_Function()

        AC_Outlet_Drawing_Create.Enabled = True
        AC_Outlet_Result_Part_Create.Enabled = True
        Message_Label.Text = "AC Outlet 간섭체크 완료"
    End Sub


    Private Sub AC_Outlet_Parameter_Control_Click(sender As Object, e As EventArgs) Handles AC_Outlet_Parameter_Control.Click
        '--------------------------------------------------------------------------------------------------- 화면 전환
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
        '--------------------------------------------------------------------------------------------------- Template_Parameter_Excel_Value 가져오기
        Call Get_Template_Parameter_Excel_Value(STW_COMP_Infomation(6, 6))
        '--------------------------------------------------------------------------------------------------- Parameter 창 Show
        ZB_Template_Parameter_Control.Show()
        Call Template_Parameter_Modify(STW_COMP_Infomation(6, 6))
        Call Folder_File_Count()
    End Sub


    Private Sub AC_Outlet_Drawing_Create_Click(sender As Object, e As EventArgs) Handles AC_Outlet_Drawing_Create.Click
        ProgressBar1.Value = 0

        Message_Label.Text = "Template02 생성중..."
        ProgressBar1.Value = 10

        '--------------------------------------------------------------------------------------------------- 화면변경
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

        '--------------------------------------------------------------------------------------------------- COMP 저장
        CATIA.ActiveDocument.Save()

        '--------------------------------------------------------------------------------------------------- Template 도면 Open
        Dim AC_Outlet_drawing_open = CATIA.Documents.NewFrom(AC_Outlet_folder_name & "\" & STW_Template_Infomation(Form_Click_Index_Value, 6))
        ProgressBar1.Value = 40

        Call CATIA_BASIC03()
        DrawingSheets.Item(1).Activate()

        '--------------------------------------------------------------------------------------------------- 도면 정보에서 Item 가져오기
        Dim Drawing_Item_Folder_Value = AC_Outlet_folder_name
        Dim Drawing_Item_Count
        For Drawing_Item_Count = 1 To 20
            If STW_Template_Infomation(Form_Click_Index_Value, 1) = STW_Drawing_Infomation(0, Drawing_Item_Count) Then
                Exit For
            End If
        Next
        '--------------------------------------------------------------------------------------------------- 도면 Part List 생성하기
        For i = 1 To STW_Drawing_Infomation_Count
            If Not STW_Drawing_Infomation(i, Drawing_Item_Count) = "" Then
                DrawingSheets.Item(STW_Drawing_Infomation(i, Drawing_Item_Count)).Activate()
                If i = 1 Then
                    '--------------------------------------------------------------------------------------------------- Block_Valve 가 있을때...
                    If Use_Block_Valve_Value = 1 Then
                        Call Template_Title_Box_Create(Drawing_Item_Folder_Value & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7), STW_Drawing_Infomation(1, 7), 1)
                        '--------------------------------------------------------------------------------------------------- Block_Valve 가 없을때...
                    Else
                        Call Template_Title_Box_Create(Drawing_Item_Folder_Value & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7), STW_Drawing_Infomation(2, 7), 1)
                    End If
                    i = i + 1
                Else
                    DrawingSheets.Item(STW_Drawing_Infomation(i, Drawing_Item_Count)).Activate()
                    Call Template_Title_Box_Create(Drawing_Item_Folder_Value & "\" & STW_Template_Infomation(Form_Click_Index_Value, 7), STW_Drawing_Infomation(i, Drawing_Item_Count), i - 1)
                End If
            End If
        Next i
        ProgressBar1.Value = 60

        On Error Resume Next
        '--------------------------------------------------------------------------------------------------- 도면 DWG_No 변경
        Drawingselection.clear
        Drawingselection.search("Name=Model*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = STW_Drawing_Infomation(1, 1)
        Next
        Drawingselection.clear
        '--------------------------------------------------------------------------------------------------- DWG_No
        Drawingselection.search("Name=DWG_No*,all")
        For i = 1 To Drawingselection.Count
            Drawingselection.Item(i).Value.Text = Template_Part_Number_Text.Text
        Next
        On Error GoTo 0

        ProgressBar1.Value = 80

        '--------------------------------------------------------------------------------------------------- 표제란 입력
        Call drawing_title_box_write()
        DrawingViews.Item("Main View").Activate()

        '--------------------------------------------------------------------------------------------------- 파일 Update
        CATIA.ActiveDocument.Update()

        '--------------------------------------------------------------------------------------------------- 도면 파일 저장
        CATIA.ActiveDocument.SaveAs(AC_Outlet_folder_name & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATDrawing")

        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        ProgressBar1.Value = 100
        Message_Label.Text = "Template02 생성완료"
    End Sub


    Private Sub AC_Outlet_Result_Part_Create_Click(sender As Object, e As EventArgs) Handles AC_Outlet_Result_Part_Create.Click
        ProgressBar1.Value = 0
        Message_Label.Text = "Template03 생성중..."

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)

        '--------------------------------------------------------------------------------------------------- BOV Hide
        selection.Clear()
        Dim visPropertySet1 = selection.VisProperties
        selection.Add(CATIA.ActiveDocument.Product.products.Item(3))
        visPropertySet1 = visPropertySet1.Parent
        visPropertySet1.SetShow(1)
        selection.Clear()
        CATIA.ActiveDocument.Save()

        '--------------------------------------------------------------------------------------------------- Product -> Part 변환
        Call Generate_CATPart_from_Product()

        '--------------------------------------------------------------------------------------------------- Geometrycal 이름에 "\" 기호 삭제
        Call Delete_Symbol()

        ProgressBar1.Value = 10

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change("CATPart", 1, 7)
        '--------------------------------------------------------------------------------------------------- PartNumber 변경
        CATIA.ActiveDocument.Product.PartNumber = Template_Part_Number_Text.Text & "_" & date_value
        publications1 = CATIA.ActiveDocument.Product.publications
        For i = 1 To Input_Output_Excel_Value_Count - 1
            '--------------------------------------------------------------------------------------------------- Publication 생성
            If Input_Output_Value(4, i) = "E" Then
                Call publication_axis_create(Input_Output_Value(1, i))
                '--------------------------------------------------------------------------------------------------- Parameter 생성
            ElseIf Input_Output_Value(4, i) = "F" Then
                Call part_parameter_create(Input_Output_Value(1, i), Input_Output_Parameter_Value(i), Input_Output_Value(2, i))
            End If
        Next
        '--------------------------------------------------------------------------------------------------- 동일한 이름의 Data가 있는지 확인
        Call Save_File_Check(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
        '--------------------------------------------------------------------------------------------------- Part File Save
        CATIA.ActiveDocument.SaveAs(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
        '--------------------------------------------------------------------------------------------------- Part Data Close
        CATIA.ActiveDocument.Close()

        '--------------------------------------------------------------------------------------------------- 도면 파일 Close
        Call CATIA_BASIC02()

        Call Active_Window_Change("CATDrawing", 1, 10)
        If Strings.Right(CATIA.ActiveDocument.Name, 10) = "CATDrawing" Then
            CATIA.ActiveDocument.Close()
        End If

        '--------------------------------------------------------------------------------------------------- Template 파일 Close
        Call CATIA_BASIC02()
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
        If CATIA.ActiveDocument.Name = STW_Template_Infomation(Form_Click_Index_Value, 5) Then
            CATIA.ActiveDocument.Close()
        End If
        ProgressBar1.Value = 30

        '--------------------------------------------------------------------------------------------------- COMP 에서 Template Data 제거
        Call CATIA_BASIC02()
        selection.Clear()
        selection.Add(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count))
        selection.Delete()
        Threading.Thread.Sleep(500)

        '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
        Call Data_Constraint_Delete(AC_Outlet_Initial_Constraint, AC_Outlet_Initial_Constraint_Count, Result_Template_Item_Value)
        '--------------------------------------------------------------------------------------------------- COMP Data에 실적 DATA 삽입
        Dim arrayOfVariantOfBSTR1(0)
        arrayOfVariantOfBSTR1(0) = sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart"
        CATIA.ActiveDocument.Product.products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
        Threading.Thread.Sleep(500)

        '--------------------------------------------------------------------------------------------------- AC_Outlet EPN 변경
        If Strings.Left(CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name, 1) <> "E" Then
            CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value
        End If

        '--------------------------------------------------------------------------------------------------- AC Outlet Data 저장
        Dim AC_Outlet_Part = Documents.Item(Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
        AC_Outlet_Part.Save()
        ProgressBar1.Value = 60

        '--------------------------------------------------------------------------------------------------- AC Outlet Constraint Delete & Create
        For i = 1 To AC_Outlet_Selection_Constraint_Count - 1
            Call Constraint_Delete(AC_Outlet_Selection_Constraint(i, 1), AC_Outlet_Selection_Constraint(i, 2), AC_Outlet_Selection_Constraint(i, 3))
        Next
        '--------------------------------------------------------------------------------------------------- BOV Constraint
        'Call Constraint_Delete("EPN0801_P02", "EPN080104_P01", 0)             'EPN 찾기.
        ProgressBar1.Value = 80
        '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
        Call O_BOM_Excel_Writing(Result_Template_Item_Value, Template_Part_Number_Text.Text & "_" & date_value)
        Call Btn_Refresh_Click()
        Threading.Thread.Sleep(500)

        '--------------------------------------------------------------------------------------------------- 파일 속성 변경 (읽기 전용)
        fso = CreateObject("Scripting.FileSystemObject")
        Dim file_Attributes_value = fso.GetFile(sDirectory_Path_Text & "\" & STW_Template_Infomation(Form_Click_Index_Value, 3) & "\" & Template_Part_Number_Text.Text & "_" & date_value & ".CATPart")
        file_Attributes_value.Attributes = 33
        On Error GoTo 0
        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()

        ProgressBar1.Value = 100
        Message_Label.Text = "Template03 생성완료"
    End Sub


    Private Sub VScroll1_Scroll(sender As Object, e As ScrollEventArgs) Handles VScroll1.Scroll
        Try
            selection.Clear()
            selection.search("Name='BLV_BOV_Rotate_Manual',all")
            If Not selection.Count = 0 Then
                selection.Item(1).Value.Value = "true"

                If Not Format(VScroll1.Value) = 0 Then
                    Count_Rotate_Text.Text = (Format(VScroll1.Value)) / 2
                Else
                    Count_Rotate_Text.Text = Format(VScroll1.Value)
                End If

                Count_Rotate_Parameter_Value.Value = Count_Rotate_Text.Text
                CATIA.ActiveDocument.Product.Update()
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("VScroll1_Scroll() 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    Private Sub VScroll2_Scroll(sender As Object, e As ScrollEventArgs) Handles VScroll2.Scroll
        Try
            selection.Clear()
            selection.search("Name='BLV_BOV_Rotate_Manual',all")
            If Not selection.Count = 0 Then
                selection.Item(1).Value.Value = "true"

                If Not Format(VScroll2.Value) = 0 Then
                    BOV_Rotate_Text.Text = (Format(VScroll2.Value)) / 2
                Else
                    BOV_Rotate_Text.Text = Format(VScroll2.Value)
                End If

                BOV_Rotate_Parameter_Value.Value = BOV_Rotate_Text.Text
                CATIA.ActiveDocument.Product.Update()
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("VScroll2_Scroll() 함수 오류!!", "", 1, 1)
        End Try
    End Sub

End Class