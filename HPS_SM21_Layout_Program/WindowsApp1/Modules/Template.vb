Imports HybridShapeTypeLib


Module Template
    '--------------------------------------------------------------------------------------------------- Input_Output_Value
    Public Input_Output_Value, Input_Output_Excel_Value_Count
    '--------------------------------------------------------------------------------------------------- Coupling 폴더 경로 정보
    Public Coupling_Repeat_Folder
    '--------------------------------------------------------------------------------------------------- Part 변환 후 Data 저장전 Parameter 생성
    Public Input_Output_Parameter_Value
    '--------------------------------------------------------------------------------------------------- Template 도면 관련 정보
    Public STW_Drawing_Infomation_Count, STW_Drawing_Infomation
    '--------------------------------------------------------------------------------------------------- Detail View 사용
    Public MyComponentInst
    '--------------------------------------------------------------------------------------------------- Parameter 수정관련 변수
    Public Template_Parameter_Idex, Template_Parameter_Fig, Template_Parameter_Name_Value, Template_Parameter_Description, Template_Parameter_Path, Template_Parameter_Excel_Count
    Public Template_Parameter_Value(200)
    Public FIG_Data_Count
    '--------------------------------------------------------------------------------------------------- 도면 Part List 가져오기
    Public template_drawing_data01(100) As String, template_drawing_data02(100) As String, template_drawing_data03(100) As String, template_drawing_data04(100) As String, template_drawing_data05(100) As String, template_drawing_data06(100) As String, template_drawing_data07(100) As String
    Public Template_Data_Folder_Name

    Private Property Translate_axis_Geom As Object



#Region " --- Input_Output_Excel_Value() : 기종_Input_Output_Sheet.xlsx 값 읽어 오기 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : sheet_name : 기종별 읽어 올 엑셀의 시트 이름
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_Excel_Value(sheet_name)
        Try
            Input_Output_Excel_Value_Count = 1
            Dim Check = True

            '--------------------------------------------------------------------------------------------------- Input_Output_Excel Open
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 2))

            Do
                '--------------------------------------------------------------------------------------------------- Input_Output_Excel_Value Count
                If Exc.Worksheets(sheet_name).Range("A" & Input_Output_Excel_Value_Count + 1).Value = "" Then        '조건이 True이면
                    Check = False
                    Exit Do
                End If

                Input_Output_Excel_Value_Count = Input_Output_Excel_Value_Count + 1
            Loop Until Check = False

            '--------------------------------------------------------------------------------------------------- 배열 동적 할당
            ReDim Input_Output_Value(0 To 4, 0 To Input_Output_Excel_Value_Count - 1)
            For i = 1 To Input_Output_Excel_Value_Count - 1
                '--------------------------------------------------------------------------------------------------- Input_Output_Excel_Value 가져오기
                Input_Output_Value(1, i) = Exc.Worksheets(sheet_name).Range("A" & i + 1).Value
                Input_Output_Value(2, i) = Exc.Worksheets(sheet_name).Range("B" & i + 1).Value
                Input_Output_Value(3, i) = Exc.Worksheets(sheet_name).Range("C" & i + 1).Value
                Input_Output_Value(4, i) = Exc.Worksheets(sheet_name).Range("F" & i + 1).Value
            Next

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(False)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL")
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Input_Output_Excel_Value 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


    Public Function Get_Template_Parameter_Excel_Value(EPN_NO_Value)
        Try
            '--------------------------------------------------------------------------------------------------- Excel Open 값 가져오기
            Exc = CreateObject("excel.application")
            'Exc.Visible = True

            sExcelFilePath = A_Main_Form.Constraint_Excel_Path_Text.Text & "\" & STW_COMP_Infomation(1, 6)
            If EXISTS_FILE_CHECK(sExcelFilePath) Then
                Exc.Workbooks.Open(sExcelFilePath)
            Else : Return False
            End If

            '--------------------------------------------------------------------------------------------------- 배열 동적 할당
            ReDim Template_Parameter_Idex(0 To 200)
            ReDim Template_Parameter_Fig(0 To 200)
            ReDim Template_Parameter_Name_Value(0 To 200)
            ReDim Template_Parameter_Description(0 To 200)
            ReDim Template_Parameter_Path(0 To 200)

            Template_Parameter_Excel_Count = 1
            Dim Check = True

            Do   'Outer loop
                '--------------------------------------------------------------------------------------------------- Template_Parameter_Fig 가져오기
                Template_Parameter_Idex(Template_Parameter_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("C" & Template_Parameter_Excel_Count + 1).Value
                '--------------------------------------------------------------------------------------------------- Template_Parameter_Fig 가져오기
                Template_Parameter_Fig(Template_Parameter_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("B" & Template_Parameter_Excel_Count + 1).Value
                '--------------------------------------------------------------------------------------------------- Template_Parameter_Name_Value 가져오기
                Template_Parameter_Name_Value(Template_Parameter_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("D" & Template_Parameter_Excel_Count + 1).Value
                '--------------------------------------------------------------------------------------------------- Template_Parameter_Description 가져오기
                Template_Parameter_Description(Template_Parameter_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("E" & Template_Parameter_Excel_Count + 1).Value
                '--------------------------------------------------------------------------------------------------- Template_Parameter_Path 가져오기
                Template_Parameter_Path(Template_Parameter_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("F" & Template_Parameter_Excel_Count + 1).Value

                Template_Parameter_Excel_Count = Template_Parameter_Excel_Count + 1

                '--------------------------------------------------------------------------------------------------- 조건 판별
                If Exc.Worksheets(EPN_NO_Value).Range("D" & Template_Parameter_Excel_Count + 1).Value = "" Then        '조건이 True이면
                    Check = False
                    Exit Do
                End If
            Loop Until Check = False

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(False)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL")

            Return True
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Get_Template_Parameter_Excel_Value 함수 오류!!", "", 1, 1)
            Return False
        End Try
    End Function


#Region " --- ZB_폼.... 검토 필요.. "
    '---------------------------------------------------------------------------------------------------
    ' Get_Template_Parameter_Excel_Value &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Template_Parameter_Modify(EPN_NO_Value)
        Dim enumValues() As Object

        Call CATIA_BASIC02()
        '--------------------------------------------------------------------------------------------------- Spread 초기화
        selection.clear
        selection.search("Name='" & EPN_NO_Value & "_Manual_Change',all")

        selection.Item(1).Value.Value = "True"
        CATIA.ActiveDocument.Product.Update()
        Threading.Thread.Sleep(1000)

        '--------------------------------------------------------------------------------------------------- Spread 초기화
        Call Clear_Spr()

        '--------------------------------------------------------------------------------------------------- ZB_Template_Parameter_Control Show
        ZB_Template_Parameter_Control.Show()
        ZB_Template_Parameter_Control.Template_Parameter_Control_EPN_Label.Text = EPN_NO_Value
        ZB_Template_Parameter_Control.Template_Parameter_Control_Lebel.Text = "Template Parameter Search ..."

        '--------------------------------------------------------------------------------------------------- Template_Spread Clear
        With ZB_Template_Parameter_Control.DataGridView1
            .Columns(0).Width = 5
            .Columns(1).Width = 5
            .Columns(2).Width = 15
            .Columns(3).Width = 7
            .Columns(4).Width = 30
            .Columns(5).Width = 40

            Call CATIA_BASIC02()

            '--------------------------------------------------------------------------------------------------- Template_Spread Clear
            .ColumnCount = 6
            .RowCount = Template_Parameter_Excel_Count - 1

            .Rows(1).Cells(0).Value = "IDEX"
            .Rows(2).Cells(0).Value = "FIG"
            .Rows(3).Cells(0).Value = "Parameter Name"
            .Rows(4).Cells(0).Value = "Value"
            .Rows(5).Cells(0).Value = "Description"
            .Rows(6).Cells(0).Value = "Parameter Path"

            For i = 1 To Template_Parameter_Excel_Count - 1
                .Rows(1).Cells(i).Value = Template_Parameter_Idex(i)
                .Rows(2).Cells(i).Value = Template_Parameter_Fig(i)
                .Rows(3).Cells(i).Value = Template_Parameter_Name_Value(i)
                .Rows(5).Cells(i).Value = Template_Parameter_Description(i)
                selection.clear
                '--------------------------------------------------------------------------------------------------- Parameter Search
                selection.search("Name='" & Template_Parameter_Name_Value(i) & "',all")
                Template_Parameter_Value(i) = selection.Item(1).Value

                '--------------------------------------------------------------------------------------------------- Single , Multi Value
                If Template_Parameter_Value(i).GetEnumerateValuesSize() - 1 = -1 Then
                    .Rows(4).Cells(i).Value = Template_Parameter_Value(i).Value
                    .RowCount = i
                    .ColumnCount = 4
                Else
                    ReDim enumValues(Template_Parameter_Value(i).GetEnumerateValuesSize() - 1)
                    Template_Parameter_Value(i).GetEnumerateValues(enumValues)
                    .RowCount = i
                    .ColumnCount = 4
                    '.CellType = 8  

                    Dim Combo_Box_Initial_Value = ""
                    For j = LBound(enumValues) To UBound(enumValues)
                        Combo_Box_Initial_Value = Combo_Box_Initial_Value & CStr(enumValues(j)) & vbTab
                    Next

                    '.TypeComboBoxList = Combo_Box_Initial_Value
                    For j = LBound(enumValues) To UBound(enumValues)
                        If Template_Parameter_Value(i).Value = CStr(enumValues(j)) Then
                            '.TypeComboBoxCurSel = j
                        End If
                    Next
                End If

                '--------------------------------------------------------------------------------------------------- Parameter Search
                .Rows(6).Cells(i).Value = Template_Parameter_Path(i)
            Next
        End With

        ZB_Template_Parameter_Control.Template_Parameter_Control_Lebel.Text = "Template Parameter Search 완료"
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' DataGridView1 초기화 &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Clear_Spr()
        ZB_Template_Parameter_Control.DataGridView1.ColumnCount = 1
        'ZB_Template_Parameter_Control.DataGridView1.Col2 = ZB_Template_Parameter_Control.DataGridView1.MaxCols
        ZB_Template_Parameter_Control.DataGridView1.RowCount = 1
        'ZB_Template_Parameter_Control.DataGridView1.Row2 = ZB_Template_Parameter_Control.DataGridView1.MaxRows
        ZB_Template_Parameter_Control.DataGridView1.Text = ""

        For i = 1 To ZB_Template_Parameter_Control.DataGridView1.ColumnCount
            ZB_Template_Parameter_Control.DataGridView1.Columns(0).Width = 10        ' 기본세로간격값
        Next i

        ZB_Template_Parameter_Control.DataGridView1.RowCount = 2
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Folder_File_Count  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Folder_File_Count()
        FIG_Data_Count = 0
        Dim retstr = ""

        '--------------------------------------------------------------------------------------------------- SMx100
        Dim FIG_Path
        If Mid(A_Main_Form.Lbl_ModelNumber.Text, 4, 1) = "1" Then
            FIG_Path = "\UI_Parameter_SMx100\FIG\"
        Else
            FIG_Path = "\UI_Parameter\FIG\"
        End If

        retstr = Dir(A_Main_Form.Constraint_Excel_Path_Text.Text & FIG_Path & ZB_Template_Parameter_Control.Template_Parameter_Control_EPN_Label.Text &
                     "\*.*", (vbDirectory + vbHidden + vbReadOnly + vbSystem))

        Do While retstr <> ""
            If retstr <> "." And retstr <> ".." Then
                FIG_Data_Count = FIG_Data_Count + 1
            End If
            retstr = Dir()
        Loop

        'lsm 190408 / ZB_Template_Parameter_Control 폼 생성시 확인 해야함
        'For i = 1 To 10
        '    ZB_Template_Parameter_Control.Fig_Show(i).Visible = False
        'Next

        'For i = 1 To FIG_Data_Count
        '    ZB_Template_Parameter_Control.Fig_Show(i).Visible = True
        'Next
    End Sub

#End Region


#Region " --- Folder_Delete() : 폴더 삭제하기 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : Folder_Value : 폴더 경로
    '---------------------------------------------------------------------------------------------------
    Public Sub Folder_Delete(Folder_Value)
        Try
            '--------------------------------------------------------------------------------------------------- Folder Copy 중복 방지
            fso = CreateObject("Scripting.FileSystemObject")
            Dim Delete_folder = fso.GetFolder(Folder_Value)
            Dim GetFolderFiles = Delete_folder.Files
            CATIA.DisplayFileAlerts = False
            For Each EachFile In GetFolderFiles
                Dim ForFileName = EachFile.Name
                If InStr(ForFileName, ".CAT") <> 0 Then
                    Dim isOpen = False
                    For i = 1 To CATIA.Documents.Count
                        If CATIA.Documents.Item(i).Name = ForFileName Then
                            isOpen = True
                        End If
                    Next i
                    If isOpen = True Then
                        CATIA.Documents.item(ForFileName).close
                    End If
                End If
            Next
            CATIA.DisplayFileAlerts = True
            Call KillProcess("EXCEL")
            Delete_folder.Name = Delete_folder.Name & "_OLD_" & Format(Now, "yymmddhhmm")

            Threading.Thread.Sleep(500)
        Catch ex As IO.DirectoryNotFoundException
            ' 없으면 무시..
        Catch ex As System.Security.SecurityException
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Folder_Delete 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- wirte_transfer_file() : 폴더의 읽기 전용 옵션 변경 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : transfer_folder_path : 폴더 경로
    '---------------------------------------------------------------------------------------------------
    Public Sub wirte_transfer_file(transfer_folder_path)
        On Error Resume Next
        '--------------------------------------------------------------------------------------------------- Folder 읽기 전용 변경
        A_Main_Form.File_numbering.Path = transfer_folder_path

        For i = 0 To A_Main_Form.File_numbering.Items.Count - 1
            Dim file_Attributes_value = fso.GetFile(transfer_folder_path & "\" & A_Main_Form.File_numbering.Items.Item(i))
            file_Attributes_value.Attributes = 32
        Next
        On Error GoTo 0
    End Sub
#End Region


#Region " --- Input_Output_Axis_Replace() : 기종_Input_Output_Sheet.xlsx의 A - Axis Replace "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Input_Value01 : 변경 할 Axis
    '             Input_Value02 : 원본 Axis
    '             Skeleton_Part : 원본 파트
    '             Copy_Geomerical_Set : ?
    '             TR_Geomerical_Set : Axis가 복사 된 Geomerical_Set
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_Axis_Replace(Input_Value01, Input_Value02, Skeleton_Part, Copy_Geomerical_Set, TR_Geomerical_Set)
        Try '---------------------------------------------------------------------------------------------------  기존 Axis 검색 backup
            'selection.search("Name='" & Input_Value01 & "',all")
            '---------------------------------------------------------------------------------------------------
            ' 생 성 일  : 20.04.23
            ' 생 성 자  : -
            ' 수 정 일  : -
            ' 수 정 자  : 김원주
            ' 수정사유  : 자냉(S)이 경우 강냉(F) Farameter값 복사 Pass           
            ' Parameter : -
            '---------------------------------------------------------------------------------------------------
            selection.Clear
            Dim MoterCoolingType_Array(5)
            MoterCoolingType_Array(0) = "TR_EPN05_P12"
            MoterCoolingType_Array(1) = "TR_EPN05_P13"
            MoterCoolingType_Array(2) = "TR_EPN02_MBFI"
            MoterCoolingType_Array(3) = "TR_EPN02_MBRI"
            MoterCoolingType_Array(4) = "TR_EPN02_MBFO"
            MoterCoolingType_Array(5) = "TR_EPN02_MBRO"
            selection.Clear
            selection.Search("name= EPN02_D10 ,all")
            Dim MoterCollingType = selection.item(1).Value.Value
            If MoterCollingType = "S" Then
                For i = 0 To 5
                    If MoterCoolingType_Array(i) = Input_Value02 Then
                        Exit Sub
                    End If

                Next

            End If
                selection.clear
                '---------------------------------------------------------------------------------------------------  기존 Axis 검색 backup
                'selection.search("Name='" & Input_Value01 & "',all")
                '---------------------------------------------------------------------------------------------------
                ' 생 성 일  : -
                ' 생 성 자  : vb6에서 사용 된 함수
                ' 수 정 일  : 19.08.22
                ' 수 정 자  : 허혜원
                ' 수정사유  : 복사할 Axis 원본 Search 로 찾지 않고 이름비교하여 찾기
                '           : 
                ' Parameter : -
                '---------------------------------------------------------------------------------------------------
                Dim PartDoc, Axis
                Dim FindItem = False
                For i = 1 To Documents.Count
                    selection.clear
                    Debug.Print(TypeName(Documents.item(i)) & " : " & Documents.item(i).Name)
                    If TypeName(Documents.item(i)) = "PartDocument" Then
                        PartDoc = Documents.item(i)
                        For j = 1 To PartDoc.Part.AxisSystems.Count
                            Axis = PartDoc.Part.AxisSystems.Item(j)
                            If Axis.Name = Input_Value01 Then
                                selection.add(Axis)
                                Exit For
                            End If
                        Next j
                    End If
                    If selection.count <> 0 Then
                        Exit For
                    End If
                Next i

                If selection.Count = 0 Then
                    selection.clear
                    selection.search("Name=' " & Input_Value01 & "',all")
                    If selection.Count <> 0 Then
                        Axis = selection.Item(1).Value
                    End If
                    'selection.search("Name='" & Input_Value01 & "',all")
                End If

                Threading.Thread.Sleep(EA_Oil_Piping_Template_Create.OilPipe_Delay_Text.Text)       'QQQ

                '--------------------------------------------------------------------------------------------------- Copy 대상이 없을때...
                If Not selection.Count = 0 Then
                    Dim Copy_Axis As Object
                    Copy_Axis = selection.Item(1).Value

                    Dim IsExists = False
                    For i = 1 To Skeleton_Part.AxisSystems.count
                        If Skeleton_Part.AxisSystems.Item(i).Name = Input_Value01 Then
                            IsExists = True
                            Exit For
                        End If
                    Next i
                    If IsExists = True Then
                        Exit Sub
                    End If
                    '--------------------------------------------------------------------------------------------------- Plane Axis Copy
                    selection.Clear()
                    selection.Add(Copy_Axis)
                    Threading.Thread.Sleep(EA_Oil_Piping_Template_Create.OilPipe_Delay_Text.Text)   'QQQ
                    selection.Copy()
                    Threading.Thread.Sleep(EA_Oil_Piping_Template_Create.OilPipe_Delay_Text.Text)   'QQQ
                    selection.Clear()
                    selection.Add(Skeleton_Part)
                    Threading.Thread.Sleep(EA_Oil_Piping_Template_Create.OilPipe_Delay_Text.Text)   'QQQ
                    selection.PasteSpecial("CATPrtResultWithOutLink")
                    Threading.Thread.Sleep(EA_Oil_Piping_Template_Create.OilPipe_Delay_Text.Text)   'QQQ
                    Skeleton_Part.Update()
                    Threading.Thread.Sleep(EA_Oil_Piping_Template_Create.OilPipe_Delay_Text.Text)   'QQQ

                    Dim Translate_axis_Geom As Object
                    Translate_axis_Geom = Skeleton_Part.HybridBodies.Item(TR_Geomerical_Set).hybridshapes

                    '--------------------------------------------------------------------------------------------------- Axis Translate 가져오기 (Axis 2개일때...)
                    ' Imports HybridShapeTypeLib / HybridShapeTranslate 참조해야 실행 됨
                    Dim Translate_axis As HybridShapeTranslate
                    Translate_axis = Translate_axis_Geom.Item(Input_Value02)

                    '--------------------------------------------------------------------------------------------------- Axis System 가져오기
                    Dim AxisSystems As Object = Skeleton_Part.AxisSystems
                    Dim axisSystem As Object = AxisSystems.Item(Input_Value01)
                    Dim reference1 As Object = Skeleton_Part.CreateReferenceFromObject(axisSystem)

                    '--------------------------------------------------------------------------------------------------- Axis Translate
                    Translate_axis.ElemToTranslate = reference1
                    Threading.Thread.Sleep(EA_Oil_Piping_Template_Create.OilPipe_Delay_Text.Text)       'QQQ
                    Skeleton_Part.Update()
                    Threading.Thread.Sleep(EA_Oil_Piping_Template_Create.OilPipe_Delay_Text.Text)       'QQQ

                Else
                    'MsgBox(Input_Value01 & " 가 존재하지 않습니다." & vbCrLf & "Axis 또는 정보를 확인해주세요.")
                End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 
            Call SHOW_MESSAGE_BOX("Input_Output_Axis_Replace(A) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Input_Output_Element_Replace() : 기종_Input_Output_Sheet.xlsx의 B - Element Replace "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Input_Value01 : 변경 할 Axis
    '             Input_Value02 : 원본 Axis
    '             Skeleton_Part : 원본 파트
    '             Copy_Geomerical_Set : 원본 Geomerical_Set
    '             TR_Geomerical_Set : Axis가 복사 된 Geomerical_Set
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_Element_Replace(Input_Value01, Input_Value02, Skeleton_Part, Copy_Geomerical_Set, TR_Geomerical_Set)
        Try
            selection.clear
            selection.search("Name='" & Input_Value01 & "',all")
            Threading.Thread.Sleep(500)

            '--------------------------------------------------------------------------------------------------- Copy 대상이 없을때...
            If Not selection.Count = 0 Then
                Dim Copy_Axis As Object

                Dim PubsCheck
                Dim selItem
                PubsCheck = False
                For i = 1 To selection.count
                    selItem = selection.Item(i).Value
                    If TypeName(selItem) = "Publication" Then
                        Copy_Axis = selItem
                        PubsCheck = True
                        Exit For
                    End If
                Next i
                If PubsCheck = False Then
                    MsgBox(Input_Value01 & " Publication이 존재하지 않습니다." & vbCrLf & "Publication 정보를 확인해주세요.")
                    Exit Sub
                End If
                '--------------------------------------------------------------------------------------------------- Plane Axis Copy
                selection.Clear()
                selection.Add(Copy_Axis)
                Threading.Thread.Sleep(200)
                selection.Copy()
                selection.Clear()
                selection.Add(Skeleton_Part)
                Threading.Thread.Sleep(200)
                selection.PasteSpecial("CATPrtResultWithOutLink")
                Threading.Thread.Sleep(500)
                Skeleton_Part.Update()
                Threading.Thread.Sleep(200)

                Dim Translate_axis_Geom As Object
                Translate_axis_Geom = Skeleton_Part.HybridBodies.Item(TR_Geomerical_Set).hybridshapes

                '--------------------------------------------------------------------------------------------------- Axis Translate 가져오기 (Axis 2개일때...)
                ' Imports HybridShapeTypeLib / HybridShapeTranslate 참조해야 실행 됨
                Dim Translate_axis As HybridShapeTranslate
                Translate_axis = Translate_axis_Geom.Item(Input_Value02)

                '--------------------------------------------------------------------------------------------------- Axis System 가져오기
                Dim Translate_axis_Geom02 As Object = Skeleton_Part.HybridBodies.Item(Copy_Geomerical_Set).hybridshapes
                Dim Result_Curve As Object = Translate_axis_Geom02.Item(Input_Value01)
                Dim reference1 As Object = Skeleton_Part.CreateReferenceFromObject(Result_Curve)

                '--------------------------------------------------------------------------------------------------- Axis Translate
                Translate_axis.ElemToTranslate = reference1
                Threading.Thread.Sleep(1000)
                Skeleton_Part.Update()
                Threading.Thread.Sleep(200)
            Else
                MsgBox(Input_Value01 & " 가 존재하지 않습니다." & vbCrLf & "Element 정보를 확인해주세요.")
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Input_Output_Element_Replace(B) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Input_Output_Initial_Parameter() : 기종_Input_Output_Sheet.xlsx의 C - Layout Parameter -> Template Parameter "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Parameter01 : Layout Parameter
    '             Parameter02 : Template Parameter
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_Initial_Parameter(Parameter01, Parameter02)
        Try
            selection.clear
            '--------------------------------------------------------------------------------------------------- EPN02_D06
            selection.search("Name='" & Parameter02 & "',all")
            Threading.Thread.Sleep(200)

            '--------------------------------------------------------------------------------------------------- Copy 대상이 없을때...
            If Not selection.Count = 0 Then
                Dim Parameter02_Value As Object
                Parameter02_Value = selection.Item(1).Value
                selection.clear
                '--------------------------------------------------------------------------------------------------- EPN05_D06
                selection.search("Name='" & Parameter01 & "',all")
                Threading.Thread.Sleep(200)

                If Not selection.Count = 0 Then
                    selection.Item(1).Value.Value = Parameter02_Value.Value
                End If
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Input_Output_Initial_Parameter(C) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Constraint_Delete() : 기종_Constraint_List.xlsx D - 읽어서 구속조건 삭제 or 추가 (constraint_value = 1 : 삭제, 0: 추가) "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : axis01_axisname : Axis 1
    '             axis02_axisname : Axis 2
    '             constraint_value : 구속 or 삭제 선택 인자
    '---------------------------------------------------------------------------------------------------
    Public Sub Constraint_Delete(axis01_axisname, axis02_axisname, constraint_value)
        On Error Resume Next        ' 엑셀에 정의 된 Axis가 없거나, 퍼블리케이션이 안만들어진 경우 넘기기 위해..
        Call CATIA_BASIC02()
        selection.Clear()

        Dim constraints = product_add.Connections("CATIAConstraints")

        If constraint_value = "1" Then
            '--------------------------------------------------------------------------------------------------- Constraint Delete
            For j = 1 To constraints.Count
                If constraints.Item(j).Name = axis02_axisname Then
                    selection.Add(constraints.Item(j))
                    Threading.Thread.Sleep(200)
                    selection.Delete()
                    Threading.Thread.Sleep(200)
                    selection.Clear()
                    Exit For
                End If
            Next
        Else
            Dim Split_DisplayName, Geometry_Name
            Dim High_Geometry_Name = Nothing
            Dim Parent, Check, parentName
            selection.clear
            '--------------------------------------------------------------------------------------------------- 1번째 Publication 찾기
            selection.search("Name='" & axis01_axisname & "',all")

            If Not selection.Count = 0 Then
                Dim Axis_Pub1_Search_Name = ""
                For j = 1 To selection.Count
                    Axis_Pub1_Search_Name = selection.Item(j).Value.Parent.Name

                    If Axis_Pub1_Search_Name = "Publications" Then
                        Pub1_Search = selection.Item(j).Value
                        Exit For
                    End If
                Next
                '--------------------------------------------------------------------------------------------------- 찾는 대상이 없으면 종료
                If Axis_Pub1_Search_Name <> "Publications" Then '200318
                    Debug.Print("No publication : 01 " & axis01_axisname)
                    Exit Sub
                End If

                Pub1_DisplayName = Nothing
                Pub1_DisplayName = Pub1_Search.Valuation.DisplayName
                '--------------------------------------------------------------------------------------------------- 1번째 이름 정하기
                Pub_reference1 = Nothing
                Pub_reference1 = product_add.CreateReferenceFromName(Pub1_DisplayName)
                '--------------------------------------------------------------------------------------------------- Publication 못 찾을 경우 1
                If Pub_reference1.Name = Nothing Then
                    Split_DisplayName = Split(Pub1_DisplayName, "!")
                    Geometry_Name = Split(Split_DisplayName(UBound(Split_DisplayName)), "/")

                    selection.Clear()
                    selection.search("Name='" & Geometry_Name(0) & "',all")
                    High_Geometry_Name = selection.Item(1).Value.Parent.Name
                    Pub_reference1 = product_add.CreateReferenceFromName(Split_DisplayName(0) & "!" & High_Geometry_Name & "/" & Split_DisplayName(1))
                End If
                selection.clear
                '--------------------------------------------------------------------------------------------------- 2번째 Publication 찾기
                selection.search("Name='" & axis02_axisname & "',all")
                If Not selection.Count = 0 Then
                    Dim Axis_Pub2_Search_Name = ""
                    For j = 1 To selection.Count
                        Axis_Pub2_Search_Name = selection.Item(j).Value.Parent.Name

                        If Axis_Pub2_Search_Name = "Publications" Then
                            Pub2_Search = selection.Item(j).Value
                            Exit For
                        End If
                    Next

                    '--------------------------------------------------------------------------------------------------- 찾는 대상이 없으면 종료
                    If Axis_Pub2_Search_Name <> "Publications" Then '200318
                        Debug.Print("No publication : 02 " & axis02_axisname)
                        Exit Sub
                    End If

                    Pub2_DisplayName = Nothing
                    Pub2_DisplayName = Pub2_Search.Valuation.DisplayName
                    '--------------------------------------------------------------------------------------------------- 2번째 이름 정하기
                    Pub_reference2 = Nothing
                    Pub_reference2 = product_add.CreateReferenceFromName(Pub2_DisplayName)
                    '--------------------------------------------------------------------------------------------------- Publication 못 찾을 경우 1
                    If Pub_reference2.Name = Nothing Then
                        Split_DisplayName = Split(Pub2_DisplayName, "!")
                        Geometry_Name = Split(Split_DisplayName(UBound(Split_DisplayName)), "/")

                        selection.Clear()
                        selection.search("Name='" & Geometry_Name(0) & "',all")

                        Pub_reference2 = product_add.CreateReferenceFromName(Split_DisplayName(0) & "!" & High_Geometry_Name & "/" & Split_DisplayName(1))
                    End If

                    '--------------------------------------------------------------------------------------------------- Publication 못가져오는 경우 원본 Item 경로를 검색하여 가져온다
                    If Pub_reference2.Name = Nothing Then
                        Split_DisplayName = Split(Pub2_DisplayName, "!")
                        Geometry_Name = Split(Split_DisplayName(UBound(Split_DisplayName)), "/")

                        High_Geometry_Name = selection.Item(1).Value.Parent.Parent.Parent.Name & "/" & selection.Item(1).Value.Parent.Name & "/" & selection.Item(1).Value.Parent.Name

                        Pub_reference2 = product_add.CreateReferenceFromName(Split_DisplayName(0) & "!" & High_Geometry_Name & "/" & Split_DisplayName(1))
                        Parent = selection.Item(1).Value
                        parentName = ""
                        Check = 0
                        Do While (Pub_reference2.Name = Nothing)
                            Parent = Parent.Parent
                            parentName = Parent.Name & "/" & parentName
                            Pub_reference2 = product_add.CreateReferenceFromName(Split_DisplayName(0) & "!" & parentName & Split_DisplayName(1))
                            Check = Check + 1

                            If Check > 100 Then
                                Exit Do
                            End If
                        Loop
                    End If

                    '--------------------------------------------------------------------------------------------------- 구속조건 생성
                    constraints = product_add.Connections("CATIAConstraints")
                    Dim constraint1 = constraints.AddBiEltCst(2, Pub_reference1, Pub_reference2)


                    '--------------------------------------------------------------------------------------------------- R26부터 발생하는 Over-Constrain 알림창 제거.
                    '--------------------------------------------------------------------------------------------------- Update Diagnosis Information 창 닫기
                    Dim Window_HWND As IntPtr = 0
                    Window_HWND = FindWindow(vbNullString, "Update Diagnosis Information: ")
                    Threading.Thread.Sleep(200)

                    If Window_HWND = 0 Then
                        Window_HWND = FindWindow(vbNullString, "Update Diagnosis: Product1")
                        Threading.Thread.Sleep(200)

                        If Window_HWND = 0 Then
                            Dim tmp_str = Split(CATIA.ActiveDocument.Name, ".")
                            Dim ActiveDocument_Name = tmp_str(0)

                            Window_HWND = FindWindow(vbNullString, "Update Diagnosis: " & ActiveDocument_Name)
                            Threading.Thread.Sleep(200)
                        End If
                    End If

                    If Not Window_HWND = 0 Then
                        Dim Button_HWND = FindWindowEx(Window_HWND, IntPtr.Zero, "Button", Nothing)
                        Threading.Thread.Sleep(200)

                        SendMessage(Button_HWND, BM_CLICK, 0, 0)
                        Threading.Thread.Sleep(200)
                    End If

                    constraint1.Name = axis02_axisname
                    CATIA.ActiveDocument.Product.Update()

                    '--------------------------------------------------------------------------------------------------- Constraint 오류시 Deactivate
                    If Not constraint1.Status = 0 Then
                        constraint1.Deactivate()
                    End If
                End If
            End If
        End If

        selection.Clear()
        On Error GoTo 0
    End Sub
#End Region


#Region " --- publication_axis_create() : 기종_Input_Output_Sheet.xlsx의 E - Output Axis "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : publication_name : 생성되어 있는 퍼플리케이션
    '---------------------------------------------------------------------------------------------------
    Public Sub publication_axis_create(publication_name)
        Try
            Dim publications1 As Object
            CATIA_BASIC02()
            publications1 = CATIA.ActiveDocument.Product.publications

            '--------------------------------------------------------------------------------------------------- product_add 재설정
            selection.Clear()
            selection.search("Name='" & publication_name & "',all")
            Threading.Thread.Sleep(200)

            If Not selection.Count = 0 Then
                Dim coupling_org_data As Object = selection.Item(1).Value
                Dim reference1 As Object = product_add.CreateReferenceFromName(product_add.Name & "/!" & coupling_org_data.Name)
                Dim publication1 As Object = publications1.Add(coupling_org_data.Name)

                publications1.SetDirect(coupling_org_data.Name, reference1)
                selection.Clear()
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("publication_axis_create(E) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- part_parameter_create() : 기종_Input_Output_Sheet.xlsx의 F - Output Parameter "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : parameter_name : 생성 할 파라미터 이름
    '             Parameter_Value : 파라미터 값
    '             parameter_type : 파라미터 타입
    '---------------------------------------------------------------------------------------------------
    Public Sub part_parameter_create(parameter_name, Parameter_Value, parameter_type)
        Try
            If Not Parameter_Value Is Nothing Then
                '--------------------------------------------------------------------------------------------------- Keyway Parameter 생성
                Dim oParameters As Object
                oParameters = CATIA.ActiveDocument.Part.Parameters

                If parameter_type = "String" Then
                    '--------------------------------------------------------------------------------------------------- String 일때...
                    oParameters.CreateString(parameter_name, CStr(Parameter_Value.Value))
                ElseIf parameter_type = "Length" Then
                    '--------------------------------------------------------------------------------------------------- Length 일때...
                    oParameters.CreateDimension(parameter_name, "LENGTH", CDbl(Parameter_Value.Value))
                ElseIf parameter_type = "Boolean" Then
                    '--------------------------------------------------------------------------------------------------- Boolean 일때...
                    If Parameter_Value.Value = "True" Then
                        oParameters.CreateString(parameter_name, "true")
                    Else
                        oParameters.CreateString(parameter_name, "false")
                    End If
                End If
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("part_parameter_create(F) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Input_Output_Insert_G_Type() : 기종_Input_Output_Sheet.xlsx의 G - 커플링BOM -> Parameter "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_Insert_G_Type()
        Dim sw = 0
        For i = 0 To Input_Output_Excel_Value_Count - 1
            For j = 0 To 10
                '--------------------------------------------------------------------------------------------------- Input_Output_Value의 값이 있을 경우만 실행.
                If Not Input_Output_Value(2, i) = "" Or Input_Output_Value(2, i) = "Empty" Then
                    '--------------------------------------------------------------------------------------------------- 두개의 값이 "MOTOR FLANGE KEYWAY"로 같으면
                    If Input_Output_Value(2, i) = coupling_create_value_name(0, j) Then
                        '--------------------------------------------------------------------------------------------------- "EPN03_D02" 파라미터의 값 변경.
                        ActiveDocument.Product.Parameters.Item(Input_Output_Value(1, i)).Value = coupling_create_value_name(1, j)
                        CATIA.ActiveDocument.Product.Update()

                        sw = 1
                        Exit For
                    End If
                End If
            Next j
            '--------------------------------------------------------------------------------------------------- Keyway 값이 입력되면 함수 종료.
            If sw = 1 Then
                sw = 0
                Exit For
            End If
        Next i
    End Sub
#End Region


#Region " --- Input_Output_Excel_to_Excel() : 기종_Input_Output_Sheet.xlsx의 H - BOM P/N -> Excel "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Excel_App : 엑셀 정보
    '             Excel_Sheet : 시트 이름
    '             Excel_Value1 : 변경 대상 구분 값
    '             Excel_Value2 : 변경 값
    '             Excel_Array1 : BOM 정보가 있는 배열
    '             For_J1 : 변경 값이 있는 배열의 구분 값
    '             For_J2 : 변경 값이 있는 배열의 구분 값
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_Excel_to_Excel(Excel_App, Excel_Sheet, Excel_Value1, Excel_Value2, Excel_Array1, For_J1, For_J2)
        'Excel_App.Visible = True
        Try
            For j = 2 To O_BOM_Part_Value_Count - 1
                If Excel_Value1 = Excel_Array1(j, For_J1) Then
                    Excel_App.Worksheets(Excel_Sheet).Range(Excel_Value2).Value = Excel_Array1(j, For_J2)
                    Exit For
                End If
            Next j
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Input_Output_Excel_to_Excel(H) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Input_Output_Excel_to_LayoutParameter() : 기종_Input_Output_Sheet.xlsx의 I - BOM P/N -> Parameter "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Selection_Value : 파라미터 정보
    '             Excel_Array1 : BOM 정보가 있는 배열
    '             For_I0 : ?
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_Excel_to_LayoutParameter(Selection_Value, Excel_Array1, For_I0, For_I1, For_I2, Excel_Array2, For_J1, For_J2)
        Try
            For j = 2 To O_BOM_Part_Value_Count - 1
                If Excel_Array1(For_I1, For_I0) = Excel_Array2(j, For_J1) Then
                    selection.clear
                    '--------------------------------------------------------------------------------------------------- Parameter 찾기
                    selection.search("Name='" & Excel_Array1(For_I2, For_I0) & "',all")
                    Threading.Thread.Sleep(200)

                    '--------------------------------------------------------------------------------------------------- 
                    If Not selection.Count = 0 Then
                        selection.Item(Selection_Value).Value.Value = Excel_Array2(j, For_J2)
                        selection.Clear()
                        Exit For
                    End If
                End If
            Next j
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Input_Output_Excel_to_LayoutParameter(U)  함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Input_Output_LayoutParameter_to_Drawing() : 기종_Input_Output_Sheet.xlsx의 J - Layout Parameter -> 도면Template "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Para_Value : 파라미터 정보
    '             Drawing_Value : 도면의 대상
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_LayoutParameter_to_Drawing(Para_Value, Drawing_Value)
        Try
            Call Active_Window_Change("Product", 1, 7)
            selection.clear
            selection.search("Name='" & Para_Value & "',all")
            Threading.Thread.Sleep(200)

            '--------------------------------------------------------------------------------------------------- 
            Dim Motor_Param As Object = Nothing
            If Not selection.Count = 0 Then
                Motor_Param = selection.Item(1).Value.Value
            End If

            Call Active_Window_Change("Drawing", 1, 7)
            selection.clear
            Drawingselection.search("Name='" & Drawing_Value & "',all")
            Threading.Thread.Sleep(200)

            '--------------------------------------------------------------------------------------------------- 
            If Not selection.Count = 0 Then
                For j = 1 To Drawingselection.Count
                    Drawingselection.Item(j).Value.Text = Motor_Param
                Next j
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Input_Output_LayoutParameter_to_Drawing(J) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Input_Output_LayoutParameter_to_Excel() : 기종_Input_Output_Sheet.xlsx의 K - Layout Parameter -> 도면엑셀 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Excel_App : 엑셀 정보
    '             Excel_Sheet : 시트 이름
    '             Para_Value : 파라미터 정보
    '             Excel_Value : 변경 할 도면의 엑셀
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_LayoutParameter_to_Excel(Excel_App, Excel_Sheet, Para_Value, Excel_Value)
        CATIA_BASIC02()

        Try
            selection.clear()
            selection.search("Name='" & Para_Value & "',all")
            Threading.Thread.Sleep(200)
            'Excel_App.Visible = True
            'Excel_App.Worksheets.Item(Excel_Sheet).Range(Excel_Value).Value = "A" '터미네이터
            If Not selection.Count = 0 Then
                Excel_App.Worksheets.Item(Excel_Sheet).Range(Excel_Value).Value = selection.Item(1).Value.Value
                If Excel_Value = "Z3" Then
                    Is_F_S = selection.Item(1).Value.Value
                End If
            ElseIf selection.Count = 0 Then
                Excel_App.Worksheets.Item(Excel_Sheet).Range(Excel_Value).Value = "N/A"
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Input_Output_LayoutParameter_to_Excel(K) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Input_Output_UI_to_LayoutParameter_L() : 기종_Input_Output_Sheet.xlsx의 L - UI에 있는 값 -> Template Parameter "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Layout_Parameter_Value : 폼 정보
    '             Excel_Value_1 : 파라미터 이름
    '             Excel_Value_2 : 컨트롤 이름
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_UI_to_LayoutParameter_L(Layout_Parameter_Value, Excel_Value_1, Excel_Value_2)
        Call CATIA_BASIC02()

        Try
            selection.Clear()
            selection.search("Name='" & Excel_Value_1 & "',all")
            Threading.Thread.Sleep(200)

            If Not selection.Count = 0 Then
                '--------------------------------------------------------------------------------------------------- 폼의 컨트롤들 체크 여부 판단
                If Excel_Value_2 = "UI_OIL_TANK_TYPE" Then
                    '--------------------------------------------------------------------------------------------------- Oil Tank KOR/CHN
                    'oControl = oForm.item("GroupBox6").Controls.item("GroupBox4").Controls.item("FRM_TankType").Controls.item("Option_KOR")

                    If Layout_Parameter_Value.Option_KOR.Checked = True Then
                        selection.Item(1).Value.Value = "KOR"
                    Else
                        selection.Item(1).Value.Value = "CHN"
                    End If
                ElseIf Excel_Value_2 = "UI_ENCLOSURE_TYPE" Then
                    '--------------------------------------------------------------------------------------------------- Base Frame
                    If Layout_Parameter_Value.UI_ENCLOSURE_TYPE.Checked = True Then
                        selection.Item(1).Value.Value = "Yes"
                    Else
                        selection.Item(1).Value.Value = "No"
                    End If
                ElseIf Excel_Value_2 = "UI_AUTOTRAP" Then
                    If Layout_Parameter_Value.Auto_Trap_Option1.Checked = True Then
                        selection.Item(1).Value.Value = "380"
                    Else
                        selection.Item(1).Value.Value = "50"
                    End If
                ElseIf Excel_Value_2 = "UI_INLET_FILTER" Then
                    '--------------------------------------------------------------------------------------------------- Enclosure
                    If Layout_Parameter_Value.UI_INLET_FILTER.Checked = True Then
                        selection.Item(1).Value.Value = "IN"
                    Else
                        selection.Item(1).Value.Value = "OUT"
                    End If
                ElseIf Excel_Value_2 = "UI_CONTROL_PANEL" Then
                    If Layout_Parameter_Value.UI_CONTROL_PANEL.Checked = True Then
                        selection.Item(1).Value.Value = Layout_Parameter_Value.UI_CONTROL_PANEL.Text
                    ElseIf Layout_Parameter_Value.UI_CONTROL_PANEL_2.Checked = True Then
                        selection.Item(1).Value.Value = Layout_Parameter_Value.UI_CONTROL_PANEL_2.Text
                    Else
                        selection.Item(1).Value.Value = Layout_Parameter_Value.UI_CONTROL_PANEL_3.Text
                    End If
                ElseIf Excel_Value_2 = "UI_SILENCER" Then
                    If Layout_Parameter_Value.UI_SILENCER.Checked = True Then
                        selection.Item(1).Value.Value = "NONE"
                    Else
                        selection.Item(1).Value.Value = "FLANGE"
                    End If
                ElseIf Excel_Value_2 = "UI_CANOPY" Then
                    If Layout_Parameter_Value.UI_CANOPY.Checked = 1 Then
                        selection.Item(1).Value.Value = "true"
                    Else
                        selection.Item(1).Value.Value = "false"
                    End If
                ElseIf Excel_Value_2 = "UI_DOOR_LOCK" Then
                    If Layout_Parameter_Value.UI_DOOR_LOCK.Checked = 1 Then
                        selection.Item(1).Value.Value = "true"
                    Else
                        selection.Item(1).Value.Value = "false"
                    End If
                ElseIf Excel_Value_2 = "UI_PANEL_OUTSIDE" Then
                    If Layout_Parameter_Value.UI_PANEL_OUTSIDE.Checked = 1 Then
                        selection.Item(1).Value.Value = "true"
                    Else
                        selection.Item(1).Value.Value = "false"
                    End If
                Else
                    '--------------------------------------------------------------------------------------------------- 타입 선택 -> 2, 3, 4
                    'oControl = oForm.item("GroupBox1").Controls.item("GroupBox2").Controls.item("UI_STAGE_COMBO")
                    selection.Item(1).Value.Value = Layout_Parameter_Value.UI_STAGE_COMBO.Text
                End If
            End If

            selection.clear()
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Input_Output_UI_to_LayoutParameter_L(L) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Input_Output_UI_to_LayoutParameter_M() : 기종_Input_Output_Sheet.xlsx의 M - UI에 VVIP 적용시 -> Template Parameter "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Layout_Parameter_Value : 폼 정보
    '             Excel_Value_1 : 파라미터 이름
    '             Excel_Value_2 : 컨트롤 이름
    '---------------------------------------------------------------------------------------------------
    Public Sub Input_Output_UI_to_LayoutParameter_M(Layout_Parameter_Value, Excel_Value_1, Excel_Value_2)
        Try
            selection.Clear()
            selection.search("Name='" & Excel_Value_1 & "',all")
            Threading.Thread.Sleep(200)

            '--------------------------------------------------------------------------------------------------- 
            If Not selection.Count = 0 Then
                For i = 0 To Layout_Parameter_Value.Controls.Count - 1
                    If Excel_Value_2 = Layout_Parameter_Value.Controls.Item(i).Name Then
                        '--------------------------------------------------------------------------------------------------- Excel Open 값 가져오기
                        If Layout_Parameter_Value.Controls.Item(i).checked = True Then
                            selection.Item(1).Value.Value = "true"
                        Else
                            selection.Item(1).Value.Value = "false"
                        End If
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Input_Output_UI_to_LayoutParameter_M(M) 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Template_Data_Get_Parameter "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Template_Data_Get_Parameter()
        Try
            Call CATIA_BASIC02()

            ReDim Input_Output_Parameter_Value(0 To 100)
            For i = 1 To Input_Output_Excel_Value_Count - 1
                If Input_Output_Value(4, i) = "F" Then
                    selection.Clear()
                    selection.search("Name='" & Input_Output_Value(1, i) & "',all")
                    Threading.Thread.Sleep(200)

                    Dim oSelCount As Integer = selection.Count
                    If Not oSelCount = 0 Then
                        Input_Output_Parameter_Value(i) = selection.Item(1).Value
                    End If
                End If
            Next
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Template_Data_Get_Parameter 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


    '---------------------------------------------------------------------------------------------------
    ' Template_Title_Box_Create  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Template_Title_Box_Create(File_Name, sheet_name, table_value)
        Try
            '-------------------------------------------------------------------------------------------------------- Excel Open
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(File_Name)

            '-------------------------------------------------------------------------------------------------------- 기종 정보 입력
            Exc.Worksheets.Item(1).Range("Z1").Value = assy_value_Name
            If Left(DrawingSheets.Name, 16) = "Coupling_Drawing" Then
                Exc.Worksheets.Item(1).Range("Z2").Value = coupling_create_value_name(1, 5)
            End If

            '------------------------------------------------------------------------------------------------------ ENCLOSURE 도면 정보 입력
            If STW_Template_Infomation(Form_Click_Index_Value, 1) = "ENCLOSURE" Then
                '--------------------------------------------------------------------------------------------------- SMx100 기종 일때...
                If Mid(A_Main_Form.Lbl_ModelNumber.Text, 4, 1) = "1" Then
                    'Call CATIA_Basic02
                    '---------------------------------------------------------------------------------------------------
                    selection.Clear()
                    selection.search("Name='EPN11_D02',all")
                    Exc.Worksheets.Item(1).Range("Z2").Value = selection.Item(1).Value.Value
                    '---------------------------------------------------------------------------------------------------
                    If GA_Enclosure_Template_Create.UI_INLET_FILTER.Checked = True Then
                        Exc.Worksheets.Item(1).Range("Z3").Value = "IN"
                    Else
                        Exc.Worksheets.Item(1).Range("Z3").Value = "OUT"
                    End If
                    '---------------------------------------------------------------------------------------------------
                    selection.Clear()
                    selection.search("Name='EPN11_D10',all")
                    Exc.Worksheets.Item(1).Range("Z4").Value = selection.Item(1).Value.Value
                    selection.Clear()
                    '---------------------------------------------------------------------------------------------------
                    If GA_Enclosure_Template_Create.UI_SILENCER.Checked = True Then
                        Exc.Worksheets.Item(1).Range("Z5").Value = "NONE"
                    Else
                        Exc.Worksheets.Item(1).Range("Z5").Value = "FLANGE"
                    End If
                Else
                    Exc.Worksheets.Item(1).Range("Z2").Value = "W840"
                    Exc.Worksheets.Item(1).Range("Z3").Value = GA_Enclosure_Template_Create.NOW_ENC_L.Text
                    Exc.Worksheets.Item(1).Range("Z4").Value = GA_Enclosure_Template_Create.UI_STAGE_COMBO.Text
                End If
            End If

            '-------------------------------------------------------------------------------------------------------- OIL PIPING Return Piping 도면 정보 입력
            If STW_Template_Infomation(Form_Click_Index_Value, 1) = "OIL RETURN PIPING" Then
                If sheet_name = "PartList" Then
                    Exc.Worksheets.Item(sheet_name).Range("Z1").Value = assy_value_Name
                    Exc.Worksheets.Item(sheet_name).Range("Z2").Value = EPN0605_MOP_Type.Value
                    Exc.Worksheets.Item(sheet_name).Range("Z3").Value = EPN0605_Filter_Type.Value
                End If
            End If

            Exc.AlertBeforeOverwriting = False      ' 저장시 덮어쓰기
            Exc.ActiveWorkbook.Save()
            Threading.Thread.Sleep(1000)

            Dim template_drawing_data_count = 1
            Dim Check = True
            'Exc.Visible = True

            Do   'Outer loop
                '--------------------------------------------------------------------------------------------------- Parameter 조건 가져오기
                template_drawing_data01(template_drawing_data_count) = Exc.Worksheets.Item(sheet_name).Range("A" & template_drawing_data_count).Value
                template_drawing_data02(template_drawing_data_count) = Exc.Worksheets.Item(sheet_name).Range("B" & template_drawing_data_count).Value
                template_drawing_data03(template_drawing_data_count) = Exc.Worksheets.Item(sheet_name).Range("C" & template_drawing_data_count).Value
                template_drawing_data04(template_drawing_data_count) = Exc.Worksheets.Item(sheet_name).Range("D" & template_drawing_data_count).Value
                template_drawing_data05(template_drawing_data_count) = Exc.Worksheets.Item(sheet_name).Range("E" & template_drawing_data_count).Value
                template_drawing_data06(template_drawing_data_count) = Exc.Worksheets.Item(sheet_name).Range("F" & template_drawing_data_count).Value
                template_drawing_data07(template_drawing_data_count) = Exc.Worksheets.Item(sheet_name).Range("G" & template_drawing_data_count).Value

                '--------------------------------------------------------------------------------------------------- ac_outlet_datum_axis 유무 판별
                If CStr(Exc.Worksheets.Item(sheet_name).Range("A" & template_drawing_data_count + 1).Value) = "" Then        '조건이 True이면
                    Check = False                                                                             '플래그 값을 False로 설정합니다.
                    Exit Do                                                                                       '내부 루프를 종료합니다.
                End If
                template_drawing_data_count = template_drawing_data_count + 1
            Loop Until Check = False   ' 외부 루프를 즉시 종료합니다.

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Save
            Exc.Quit()

            Call KillProcess("EXCEL")

            DrawingSheets.Item(sheet_name).Activate()
            Threading.Thread.Sleep(300)

            '--------------------------------------------------------------------------------------------------- Part List Box 가져오기
            Dim drawing_title_detail_open = CATIA.Documents.Open(Application.StartupPath & "\title_block_sample.CATDrawing")
            Threading.Thread.Sleep(500)

            Dim note_detail_sheet = drawing_title_detail_open.sheets.Item("title_block")
            Dim MyComponentRef = note_detail_sheet.views.Item("Main_Table01")
            MyComponentInst = DrawingSheets.Item(sheet_name).views.Item(1).Components.Add(MyComponentRef, 0, 0)

            Call detail_view_explode(MyComponentInst)

            'If template_drawing_data_count <36 Then '16.02.24 / 이연중 / 도면 파트리스트 카운트 36에서 31로 변경 / 31개 이상의 경우 왼쪽에 파트리스트 하나 더 생성 / 31 : 타이틀 행을 포함한 갯수 31
            If template_drawing_data_count <31 Then
                '--------------------------------------------------------------------------------------------------- Detail File Colse
                                                            drawing_title_detail_open.Close()
                Threading.Thread.Sleep(300)

                '--------------------------------------------------------------------------------------------------- Table 찾기
                Dim Main_Table01 As Object = Nothing
                Drawingselection.Clear()
                Drawingselection.search("Name='Main_Table01',all")
                If Not Drawingselection.Count = 0 Then
                    Main_Table01 = Drawingselection.Item(table_value).Value
                    Main_Table01.AnchorPoint = 2
                    Main_Table01.x = 244.074 / DrawingSheets.Item(sheet_name).scale2
                    Main_Table01.y = 67.095 / DrawingSheets.Item(sheet_name).scale2
                End If

                '--------------------------------------------------------------------------------------------------- Table 갯수 맞추기
                If Main_Table01 IsNot Nothing Then
                    For j = 1 To 31 - template_drawing_data_count
                        Main_Table01.RemoveRow(1)
                    Next
                End If

                '--------------------------------------------------------------------------------------------------- Table에 값 입력
                For i = 1 To template_drawing_data_count - 1
                    Try
                        Main_Table01.SetCellString(i, 1, CStr(template_drawing_data01(i)))
                        Main_Table01.SetCellString(i, 2, CStr(template_drawing_data02(i)))
                        Main_Table01.SetCellString(i, 3, CStr(template_drawing_data03(i)))
                        Main_Table01.SetCellString(i, 4, CStr(template_drawing_data04(i)))
                        Main_Table01.SetCellString(i, 5, CStr(template_drawing_data05(i)))
                        Main_Table01.SetCellString(i, 6, CStr(template_drawing_data06(i)))
                        '--------------------------------------------------------------------------------------------------- Table  7번째 칼럼 컨트롤 제거
                        'Main_Table01.SetCellString(i, 7, CStr(template_drawing_data07(i)))
                    Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                        '컬럼 개수가 다를 수 있어, 무시하고 진행
                    End Try
                Next
            Else
                DrawingSheets.Item(sheet_name).Activate()
                MyComponentRef = note_detail_sheet.views.Item("Main_Table02")
                MyComponentInst = DrawingSheets.Item(sheet_name).views.Item(1).Components.Add(MyComponentRef, 0, 0)

                Call detail_view_explode(MyComponentInst)
                '--------------------------------------------------------------------------------------------------- Detail File Colse
                drawing_title_detail_open.Close()
                Threading.Thread.Sleep(500)
                Drawingselection.clear
                Drawingselection.search("Name='Main_Table01',all")
                Dim Main_Table01 = Drawingselection.Item(table_value).Value
                Main_Table01.AnchorPoint = 2
                Main_Table01.x = 244.074
                Main_Table01.y = 67.095
                '--------------------------------------------------------------------------------------------------- Table에 값 입력
                For i = 1 To 30
                    Try
                        Main_Table01.SetCellString(31 - i, 1, CStr(template_drawing_data01(template_drawing_data_count - i)))
                        Main_Table01.SetCellString(31 - i, 2, CStr(template_drawing_data02(template_drawing_data_count - i)))
                        Main_Table01.SetCellString(31 - i, 3, CStr(template_drawing_data03(template_drawing_data_count - i)))
                        Main_Table01.SetCellString(31 - i, 4, CStr(template_drawing_data04(template_drawing_data_count - i)))
                        Main_Table01.SetCellString(31 - i, 5, CStr(template_drawing_data05(template_drawing_data_count - i)))
                        Main_Table01.SetCellString(31 - i, 6, CStr(template_drawing_data06(template_drawing_data_count - i)))
                        Main_Table01.SetCellString(31 - i, 7, CStr(template_drawing_data07(template_drawing_data_count - i)))
                    Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                        '컬럼 개수가 다를 수 있어, 무시하고 진행
                    End Try
                Next
                Drawingselection.Clear()
                '--------------------------------------------------------------------------------------------------- Detail File Colse
                Drawingselection.search("Name='Main_Table02',all")
                Dim Main_Table02 = Drawingselection.Item(table_value).Value
                Main_Table02.AnchorPoint = 8
                Main_Table01.x = 244.074
                Main_Table01.y = 67.095

                '--------------------------------------------------------------------------------------------------- Table 갯수 맞추기
                For j = 1 To 30 - (template_drawing_data_count - 31)
                    Main_Table02.RemoveRow(1)
                Next

                '--------------------------------------------------------------------------------------------------- Table에 값 입력
                For i = 1 To template_drawing_data_count - 31
                    Try
                        Main_Table02.SetCellString(template_drawing_data_count - 30 - i, 1, CStr(template_drawing_data01(template_drawing_data_count - (30 + i))))
                        Main_Table02.SetCellString(template_drawing_data_count - 30 - i, 2, CStr(template_drawing_data02(template_drawing_data_count - (30 + i))))
                        Main_Table02.SetCellString(template_drawing_data_count - 30 - i, 3, CStr(template_drawing_data03(template_drawing_data_count - (30 + i))))
                        Main_Table02.SetCellString(template_drawing_data_count - 30 - i, 4, CStr(template_drawing_data04(template_drawing_data_count - (30 + i))))
                        Main_Table02.SetCellString(template_drawing_data_count - 30 - i, 5, CStr(template_drawing_data05(template_drawing_data_count - (30 + i))))
                        Main_Table02.SetCellString(template_drawing_data_count - 30 - i, 6, CStr(template_drawing_data06(template_drawing_data_count - (30 + i))))
                        Main_Table02.SetCellString(template_drawing_data_count - 30 - i, 7, CStr(template_drawing_data07(template_drawing_data_count - (30 + i))))
                    Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                        '컬럼 개수가 다를 수 있어, 무시하고 진행
                    End Try
                Next
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Template_Title_Box_Create 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' detail_view_explode  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub detail_view_explode(MyComponentInst)
        On Error Resume Next
        '--------------------------------------------------------------------------------------------------- Detail View Element 요소 분해
        MyComponentInst.Explode()

        '--------------------------------------------------------------------------------------------------- Detail View 지우기
        Drawingselection.Clear()
        'Drawingselection.search("Name='Main_Table01',all")
        Drawingselection.Add(MyComponentInst)
        Drawingselection.Delete()
        Drawingselection.Clear()
        On Error GoTo 0
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' drawing_title_box_write &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub drawing_title_box_write()
        Try
            Call CATIA_BASIC03()

            '--------------------------------------------------------------------------------------------------- Excel Open
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(Application.StartupPath & "\" & "Draw_Title.xlsx")

            '--------------------------------------------------------------------------------------------------- Excel 정보 가져오기
            Dim get_drawn_text(10), get_drawn_name_text(10), get_drawn_date_text(10)
            For i = 1 To 6
                get_drawn_text(i) = Exc.Worksheets("Sheet1").Range("A" & i).Value
                get_drawn_name_text(i) = Exc.Worksheets("Sheet1").Range("B" & i).Value
                get_drawn_date_text(i) = Exc.Worksheets("Sheet1").Range("C" & i).Value
            Next

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(True)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL")

            '--------------------------------------------------------------------------------------------------- text value
            For i = 1 To 6
                Drawingselection.clear
                Drawingselection.search("Name=drawn" & i & "*,all")
                If Drawingselection.Count <> 0 Then
                    For j = 1 To Drawingselection.Count
                        Drawingselection.Item(j).Value.Text = get_drawn_text(i)
                    Next
                End If
                Drawingselection.clear
                Drawingselection.search("Name=drawn_name" & i & "*,all")
                If Drawingselection.Count <> 0 Then
                    For j = 1 To Drawingselection.Count
                        Drawingselection.Item(j).Value.Text = get_drawn_name_text(i)
                    Next
                End If
                Drawingselection.clear
                Drawingselection.search("Name=drawn_date" & i & "*,all")
                If Drawingselection.Count <> 0 Then
                    For j = 1 To Drawingselection.Count
                        Drawingselection.Item(j).Value.Text = get_drawn_date_text(i)
                    Next
                End If
            Next
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("drawing_title_box_write 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Delete_Symbol : Geometrycal 이름에 "\" 기호를 삭제한다.
    '---------------------------------------------------------------------------------------------------
    Public Sub Delete_Symbol()
        Try
            Call CATIA_BASIC02()
            selection.clear
            selection.search("Name='*\*',all")
            For i = 1 To selection.Count
                If selection.Item(i).Type = "HybridBody" Or selection.Item(i).Type = "Body" Or selection.Item(i).Type = "Add" Or selection.Item(i).Type = "Assemble" Then
                    Dim Geometrycal_Name_Split As Object
                    Geometrycal_Name_Split = Split(selection.Item(i).Value.Name, "\")

                    Dim TEMP_Name As Object = Nothing
                    For k = 0 To UBound(Geometrycal_Name_Split)
                        If k = 0 Then
                            TEMP_Name = Geometrycal_Name_Split(k)
                        Else
                            TEMP_Name = TEMP_Name & Geometrycal_Name_Split(k)
                        End If
                    Next k

                    selection.Item(i).Value.Name = TEMP_Name
                End If
            Next i
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Delete_Symbol 함수 오류!!", "", 1, 1)
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Save_File_Check  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Save_File_Check(File_Path)
        Try
            Dim Delete_File = Nothing
            Delete_File = fso.GetFile(File_Path)

            '--------------------------------------------------------------------------------------------------- Folder Copy 중복 방지
            If Not Delete_File Is Nothing Then
                Dim Delete_File_Len = Len(Delete_File.Name)
                Delete_File.Name = Left(Delete_File.Name, Delete_File_Len - 8) & "-OLD_" & Format(Now, "yymmddhhmm") & ".CATPart"
                Threading.Thread.Sleep(500)
            End If
        Catch ex As IO.FileNotFoundException

        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Save_File_Check 함수 오류!!", "", 1, 1)
        End Try
    End Sub


#Region " --- VVIP_Option() : VVIP 옵션에 따라 파라미터에 true / false 값 입력 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.14
    ' 수 정 자  : 이상민
    ' 수정사유  : 매개변수 변경 (4번 : Parameter_Value01, 5번 : Parameter_Value02 삭제)
    ' Parameter : VVIP_Option : UI에 VVIP 옵션 체크 값
    '             VVIP_Option_Check : UI에 VVIP 옵션 중 선택 값
    '             Parameter01 : 변경 할 파라미터 값
    '---------------------------------------------------------------------------------------------------
    Public Sub VVIP_Option(VVIP_Option, VVIP_Option_Check, Parameter01)
        Call CATIA_BASIC02()
        selection.Clear()

        Try
            '--------------------------------------------------------------------------------------------------- 파라미터 검색
            Dim Parameter01_Value As Object
            selection.clear
            selection.search("Name='" & Parameter01 & "',all")
            If selection.Count = 0 Then
                Exit Sub
            Else
                Parameter01_Value = selection.Item(1).Value
                selection.Clear()
            End If

            '--------------------------------------------------------------------------------------------------- VVIP Option 적용
            If VVIP_Option.Checked = True Then
                If VVIP_Option_Check.Checked = True Then
                    Parameter01_Value.Value = "true"
                Else
                    Parameter01_Value.Value = "false"
                End If
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("VVIP_Option() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Parameter_Error_Message() : VVIP 옵션에 따라 파라미터에 true / false 값 입력 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.14
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : Min_Value : 파라미터에 입력 할 수 있는 최소 값
    '             Max_Value : 파라미터에 입력 할 수 있는 최대 값
    '             Text_Value : UI에 있는 텍스트 박스
    '             Parameter_Value : 카티아에 파라미터
    '---------------------------------------------------------------------------------------------------
    Public Sub Parameter_Error_Message(Min_Value As Double, Max_Value As Double, Text_Value As TextBox, Parameter_Value As Object)
        Try
            If CDbl(Text_Value.Text) < Min_Value Then
                Call SHOW_MESSAGE_BOX(Min_Value & " ~ " & Max_Value & " 사이의 값을 입력하세요.", "", 1, 1)
                Text_Value.Text = CDbl(Parameter_Value.Value)
            ElseIf CDbl(Text_Value.Text) > Max_Value Then
                Call SHOW_MESSAGE_BOX(Min_Value & " ~ " & Max_Value & " 사이의 값을 입력하세요.", "", 1, 1)
                Text_Value.Text = Parameter_Value.Value
            Else
                Parameter_Value.Value = Text_Value.Text
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
        End Try
    End Sub
#End Region




End Module
