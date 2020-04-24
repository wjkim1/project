Imports INFITF
Imports MECMOD
Imports HybridShapeTypeLib


Module Result_Data
    '--------------------------------------------------------------------------------------------------- 선정 때 Open 된 파트의 개수
    Public FileListCounter As Integer
    Public axis_axis_measure
    '--------------------------------------------------------------------------------------------------- Result_Data_Value 정의
    Public Result_Data_Value
    '--------------------------------------------------------------------------------------------------- Result_Data 수량 정의
    Public Result_Check_Count As Integer
    '--------------------------------------------------------------------------------------------------- get_parameter_first_value
    Public get_parameter_first_value(20), get_parameter_first_count, get_parameter_second_value(20), get_parameter_second_count, data_open_count
    '--------------------------------------------------------------------------------------------------- Result_Data Version
    Public Data_Version_Infomation, Data_Version_Infomation_Count
    '--------------------------------------------------------------------------------------------------- Result_Data 내림차순 정렬
    Public Result_Data_Value_Number


#Region " --- DataGridView_Inicial_Setting() : DataGridView 초기 세팅, 컬럼의 해더, 넓이 값 등을 설정 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.09
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : DataGridView_Value : DataGridView
    '             Loop_Count : 작업 개수
    '             Parameter_Value : 엑셀에서 읽어 온 배열
    '---------------------------------------------------------------------------------------------------
    Public Sub DataGridView_Inicial_Setting(DataGridView_Value As DataGridView, Loop_Count As Object, Parameter_Value As Object)
        Try
            DataGridView_Value.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            DataGridView_Value.RowHeadersVisible = False
            DataGridView_Value.Columns(1).HeaderText = "No"
            DataGridView_Value.Columns(2).HeaderText = "이름"
            DataGridView_Value.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView_Value.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView_Value.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView_Value.Columns(0).Width = 20
            DataGridView_Value.Columns(1).Width = 30
            DataGridView_Value.Columns(2).Width = 100
            DataGridView_Value.ColumnCount = 3

            Dim DataGridView_Value_Count = 3
            '--------------------------------------------------------------------------------------------------- 수량에 맞게 Loop
            For i = 1 To Loop_Count - 1
                '--------------------------------------------------------------------------------------------------- 거리 비교 구문인지 확인
                If Parameter_Value(i, 8) = 0 Then
                    '--------------------------------------------------------------------------------------------------- X 값 참일때...
                    If Parameter_Value(i, 5) = 1 Then
                        DataGridView_Value_Count = DataGridView_Value_Count + 1
                        DataGridView_Value.ColumnCount = DataGridView_Value_Count
                        DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderText = Parameter_Value(i, 2) & "_X"
                        DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    End If

                    '--------------------------------------------------------------------------------------------------- Y 값 참일때...
                    If Parameter_Value(i, 6) = 1 Then
                        DataGridView_Value_Count = DataGridView_Value_Count + 1
                        DataGridView_Value.ColumnCount = DataGridView_Value_Count
                        DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderText = Parameter_Value(i, 2) & "_Y"
                        DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    End If

                    '--------------------------------------------------------------------------------------------------- Z 값 참일때...
                    If Parameter_Value(i, 7) = 1 Then
                        DataGridView_Value_Count = DataGridView_Value_Count + 1
                        DataGridView_Value.ColumnCount = DataGridView_Value_Count
                        DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderText = Parameter_Value(i, 2) & "_Z"
                        DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    End If

                    '--------------------------------------------------------------------------------------------------- 새로 만든 컬럼의 텍스트 정렬방식
                    DataGridView_Value.Columns(DataGridView_Value_Count - 1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    DataGridView_Value.Columns(DataGridView_Value_Count - 1).Width = 120
                Else
                    '--------------------------------------------------------------------------------------------------- 값 비교 구문인지 확인
                    DataGridView_Value_Count = DataGridView_Value_Count + 1
                    DataGridView_Value.ColumnCount = DataGridView_Value_Count

                    '--------------------------------------------------------------------------------------------------- 새로 만든 컬럼의 텍스트 정렬방식
                    DataGridView_Value.Columns(DataGridView_Value_Count - 1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                    If Right(Parameter_Value(i, 2), 19) = "Design_Rule_Version" Or Parameter_Value(i, 2) = "EPN04_DRV" Then
                        DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderText = Parameter_Value(i, 2)
                        DataGridView_Value.Columns(DataGridView_Value_Count - 1).Width = 230
                        DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    Else
                        If DataGridView_Value.Parent.Text = "Base Frame Create" Then
                            DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderText = Parameter_Value(i, 2)
                            DataGridView_Value.Columns(DataGridView_Value_Count - 1).Width = 150
                            DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                        ElseIf DataGridView_Value.Parent.Text = "Simple Base Create" Then
                            If Parameter_Value(i, 2) = "오일탱크 크기" Then
                                DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderText = Parameter_Value(i, 2)
                                DataGridView_Value.Columns(DataGridView_Value_Count - 1).Width = 150
                                DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                            Else
                                DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderText = Parameter_Value(i, 2)
                                DataGridView_Value.Columns(DataGridView_Value_Count - 1).Width = 190
                                DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                            End If
                        Else
                            DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderText = Parameter_Value(i, 2)
                            DataGridView_Value.Columns(DataGridView_Value_Count - 1).Width = 110
                            DataGridView_Value.Columns(DataGridView_Value_Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                        End If
                    End If
                End If
            Next i
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("DataGridView_Inicial_Setting() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- axis_distance() : Axis 간 거리를 측정 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.09
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : axis01_axisname : 1번 축
    '             axis02_axisname : 2번 축
    '---------------------------------------------------------------------------------------------------
    Public Sub axis_distance(axis01_axisname, axis02_axisname)
        Dim iNum As Integer
        Dim axis01_center_point(2), axis02_center_point(2)

        axis_axis_measure = 0       ' 두개의 축이 없어 측정이 안될 때 값 입력을 위해..

        Try
            Call CATIA_BASIC02()
            '--------------------------------------------------------------------------------------------------- 새로운 Part 생성
            publications1 = product_add.publications
            constraints1 = product_add.Connections("CATIAConstraints")
            selection = CATIA.ActiveDocument.selection
            selection.clear
            '--------------------------------------------------------------------------------------------------- axis01 value 가져오기
            selection.search("Name='" & axis01_axisname & "',all")

            Dim geom_set_selection = Nothing
            Dim geom_set_selection02 = Nothing
            Dim publication1 = Nothing, publication2 = Nothing
            If Not selection.Count = 0 Then
                Dim axis01_distance_value_selection = selection.Item(1).Value
                '---------------------------------------------------------------------------------------------------------------- Part Name 찾기
                Dim axis01_part = axis01_distance_value_selection.Parent.Parent
                For i = 1 To product_add.products.Count
                    If product_add.products.Item(i).PartNumber = axis01_part.Name Then
                        iNum = i
                        Exit For
                    End If
                Next

                For j = 1 To axis01_part.AxisSystems.Count
                    If axis01_part.AxisSystems.Item(j).Name = axis01_axisname Then
                        '--------------------------------------------------------------------------------------------------- axis01 value center point 생성하기
                        axis01_part = axis01_distance_value_selection.Parent.Parent

                        Dim axis01_axissystems As AxisSystems = axis01_part.AxisSystems
                        Dim axis01_axissystem As AxisSystem = axis01_axissystems.Item(axis01_distance_value_selection.Name)
                        axis01_axissystem.IsCurrent = True
                        axis01_axissystem.GetOrigin(axis01_center_point)

                        '--------------------------------------------------------------------------------------------------- 새로운 Geom.set 생성
                        geom_set_selection = axis01_part.HybridBodies.Add

                        '---------------------------------------------------------------------------------------------------------------- axis01_center_point_create 생성
                        Dim Axis_Reference1 As Reference = axis01_part.CreateReferenceFromObject(axis01_part.AxisSystems.Item(j))

                        Dim axis01_center_point_create As HybridShapePointCoord = axis01_part.hybridShapeFactory.AddNewPointCoord(0, 0, 0)

                        geom_set_selection.AppendHybridShape(axis01_center_point_create)
                        axis01_center_point_create.RefAxisSystem = Axis_Reference1
                        Axis_Reference1 = axis01_center_point_create.RefAxisSystem
                        axis01_center_point_create.Name = "axis01_center_point"

                        '---------------------------------------------------------------------------------------------------------------- axis01_center_point_publication 생성
                        Dim reference1 As Reference = product_add.CreateReferenceFromName(product_add.Name & "/" & product_add.products.Item(iNum).Name & "/!axis01_center_point")

                        publication1 = publications1.Add("axis01_center_point")
                        publications1.SetDirect("axis01_center_point", reference1)
                    End If
                Next

                '--------------------------------------------------------------------------------------------------- axis02 value 가져오기
                selection.Clear()
                selection.search("Name='" & axis02_axisname & "',all")

                If Not selection.Count = 0 Then
                    Dim axis02_distance_value_selection = selection.Item(1).Value
                    '---------------------------------------------------------------------------------------------------------------- Part Name 찾기
                    Dim axis02_part = axis02_distance_value_selection.Parent.Parent
                    For i = 1 To product_add.products.Count
                        If product_add.products.Item(i).PartNumber = axis02_part.Name Then
                            iNum = i
                            Exit For
                        End If
                    Next

                    For j = 1 To axis02_part.AxisSystems.Count
                        If axis02_part.AxisSystems.Item(j).Name = axis02_axisname Then
                            '--------------------------------------------------------------------------------------------------- axis02 value center point 생성하기
                            axis02_part = axis02_distance_value_selection.Parent.Parent

                            Dim axis02_axissystems As AxisSystems = axis02_part.AxisSystems
                            Dim axis02_axissystem As AxisSystem = axis02_axissystems.Item(axis02_distance_value_selection.Name)
                            axis02_axissystem.IsCurrent = True
                            axis02_axissystem.GetOrigin(axis02_center_point)

                            '--------------------------------------------------------------------------------------------------- 새로운 Geom.set 생성
                            geom_set_selection02 = axis02_part.HybridBodies.Add

                            '---------------------------------------------------------------------------------------------------------------- axis02_center_point_create 생성
                            Dim Axis_Reference2 As Reference = axis02_part.CreateReferenceFromObject(axis02_part.AxisSystems.Item(j))

                            Dim axis02_center_point_create As HybridShapePointCoord = axis02_part.hybridShapeFactory.AddNewPointCoord(0, 0, 0)

                            geom_set_selection02.AppendHybridShape(axis02_center_point_create)
                            axis02_center_point_create.RefAxisSystem = Axis_Reference2
                            Axis_Reference2 = axis02_center_point_create.RefAxisSystem
                            axis02_center_point_create.Name = "axis02_center_point"

                            axis02_part.Update()
                            product_add.Update()

                            '---------------------------------------------------------------------------------------------------------------- axis02_center_point_publication 생성
                            Dim reference2 As Reference = product_add.CreateReferenceFromName(product_add.Name & "/" & product_add.products.Item(iNum).Name & "/!axis02_center_point")

                            publication2 = publications1.Add("axis02_center_point")
                            publications1.SetDirect("axis02_center_point", reference2)
                        End If
                    Next
                    product_add.Update()

                    '---------------------------------------------------------------------------------------------------  publication DisplayName 가져오기
                    Pub1_DisplayName = publication1.Valuation.DisplayName
                    Pub2_DisplayName = publication2.Valuation.DisplayName
                    Pub_reference1 = product_add.CreateReferenceFromName(Pub1_DisplayName)
                    Pub_reference2 = product_add.CreateReferenceFromName(Pub2_DisplayName)

                    '--------------------------------------------------------------------------------------------------- Length Measure
                    Dim TheSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
                    Dim TheMeasurable = TheSPAWorkbench.GetMeasurable(Pub_reference1)
                    axis_axis_measure = Math.Round(TheMeasurable.GetMinimumDistance(Pub_reference2), 1)

                    '--------------------------------------------------------------------------------------------------- Data Delete
                    selection.Clear()
                    publications1.Remove("axis01_center_point")
                    publications1.Remove("axis02_center_point")
                    selection.Add(geom_set_selection)
                    selection.Add(geom_set_selection02)
                    selection.Delete()
                End If
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("axis_distance 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Clash_Check_Folder_Create() : 간섭체크 결과가 저장 될 폴더 생성 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.10
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Clash_Check_Folder_Create()      ' Folder Copy 중복 방지를 위해 실행
        Dim Delete_folder = Nothing
        Try
            Delete_folder = fso.getfolder(sOBomPath & "\CLASH_CHECK")
            If Delete_folder Is Nothing Then
                fso.CreateFolder(sOBomPath & "\CLASH_CHECK")
                Threading.Thread.Sleep(300)
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Error : 간섭체크 폴더 생성 함수", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Data_Parameter_Comparison() : 선정시 만들어진 템플릿 데이터 열어 맞는지 비교 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.02
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Data_Parameter_Comparison(Data_Search_Value, Data_Search_Directory, Form_Value, EPN_NO_Value, Selection_Parameter, Initial_Parameter, Selection_Parameter_Count)
        Try
            '--------------------------------------------------------------------------------------------------- 비교 조건에 만족하는 Data 수량 Count
            Result_Check_Count = 1
            FileListCounter = 0
            Form_Value.Data_Path_List.items.Clear()

            '--------------------------------------------------------------------------------------------------- File 이름 가져오기
            A_Main_Form.File_numbering.Path = Data_Search_Directory
            ReDim Result_Data_Value(0 To A_Main_Form.File_numbering.Items.Count + 1)
            Dim FName = Dir(Data_Search_Value, vbNormal)  '목록을 생성해서 파일명을 가져옵니다.

            Do While FName <> ""   ' 루프(loop)를 시작합니다.
                Call CATIA_BASIC02()

                Dim Result_File_Open = Nothing
                If Not FName = "" Then
                    Result_File_Open = CATIA.Documents.Open(Data_Search_Directory & FName)
                    Result_Data_Value(Result_Check_Count) = FName
                    Call Put_Parameter_Data_Comparison(Selection_Parameter, Initial_Parameter, Selection_Parameter_Count, Form_Value, Data_Search_Directory)
                End If
                FileListCounter = FileListCounter + 1
                FName = Dir()    ' 다음 항목을 읽어들입니다.

                Result_File_Open.Close()
            Loop

            '--------------------------------------------------------------------------------------------------- 일치하는 Data가 없을때...
            If Result_Check_Count = 1 Then
                Call SHOW_MESSAGE_BOX("일치하는 Data가 없습니다.", "", 1, 1)
                Form_Value.Message_Label.text = "Data 선정이 완료 되었습니다."
                Form_Value.DataGridView1.RowCount = 0       ' 선정된 데이터 없으면 그리드뷰의 row 개수 조정
            Else
                Dim temp01

                ReDim Result_Data_Value_Number(0 To Form_Value.Data_Path_List.Items.Count, 0 To Form_Value.DataGridView1.ColumnCount - 1)
                ReDim temp01(0 To Form_Value.DataGridView1.ColumnCount)

                ' 190426 그리드 정렬 함수 생각하기.. lsm
                ''--------------------------------------------------------------------------------------------------- 내림차순 정리
                'For i = 1 To Form_Value.Data_Path_List.Items.Count
                '    For j = 1 To Form_Value.DataGridView1.ColumnCount - 1
                '        Result_Data_Value_Number(i, j) = Form_Value.DataGridView1.Rows(i).Cells(j + 1).Value
                '    Next
                'Next

                'If Form_Value.Data_Path_List.Items.Count > 1 Then
                '    '--------------------------------------------------------------------------------------------------- 내림차순 정리하기
                '    For i = 1 To Form_Value.Data_Path_List.Items.Count - 2
                '        For j = i + 1 To Form_Value.Data_Path_List.Items.Count - 1
                '            If UCase(Result_Data_Value_Number(i, 1)) > UCase(Result_Data_Value_Number(j, 1)) Then
                '                For k = 1 To Form_Value.DataGridView1.ColumnCount - 1
                '                    temp01(k) = Result_Data_Value_Number(i, k)
                '                    Result_Data_Value_Number(i, k) = Result_Data_Value_Number(j, k)
                '                    Result_Data_Value_Number(j, k) = temp01(k)
                '                Next
                '            End If
                '        Next
                '    Next
                'End If

                'For i = 1 To Form_Value.Data_Path_List.Items.Count
                '    For j = 1 To Form_Value.DataGridView1.ColumnCount - 1
                '        Form_Value.DataGridView1.Rows(i).Cells(j + 1).Value = Result_Data_Value_Number(i, j)
                '    Next
                'Next

                Call CATIA_BASIC02()

                '--------------------------------------------------------------------------------------------------- Data Loading
                Dim arrayOfVariantOfBSTR1(0)
                arrayOfVariantOfBSTR1(0) = Form_Value.Data_Path_List.Items(0) & Form_Value.DataGridView1.Rows(0).Cells(2).Value

                CATIA.ActiveDocument.Product.products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

                '--------------------------------------------------------------------------------------------------- Data Loading 안됐을때...
                Dim Result_Data_Value_Name = "Not_Value"
                Result_Data_Value = CATIA.Documents.Item(Form_Value.DataGridView1.Rows(0).Cells(2).Value)
                Result_Data_Value_Name = CATIA.Documents.Item(Form_Value.DataGridView1.Rows(0).Cells(2).Value).Name
                If Result_Data_Value_Name = "Not_Value" Then
                    Call SHOW_MESSAGE_BOX("Data Open 실패. 실적 Data 정보를 확인하세요", "", 1, 1)
                Else
                    CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value
                End If
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Data_Parameter_Comparison() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Data_Parameter_Comparison_Oil_Piping() : 선정시 만들어진 템플릿 데이터 열어 맞는지 비교 (Oil Piping Range 범위 검색) "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.20
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Data_Parameter_Comparison_Oil_Piping(Data_Search_Value, Data_Search_Directory, Form_Value, EPN_NO_Value, Selection_Parameter, Initial_Parameter, Selection_Parameter_Count)
        Try
            '--------------------------------------------------------------------------------------------------- 비교 조건에 만족하는 Data 수량 Count
            Result_Check_Count = 1
            FileListCounter = 0
            Form_Value.Data_Path_List.Clear()

            '--------------------------------------------------------------------------------------------------- File 이름 가져오기
            A_Main_Form.File_numbering.Path = Data_Search_Directory
            ReDim Result_Data_Value(0 To A_Main_Form.File_numbering.Items.Count + 1)

            Dim Result_File_Open = Nothing
            Dim FName = Dir(Data_Search_Value, vbNormal)  '목록을 생성해서 파일명을 가져옵니다.
            Do While FName <> ""   ' 루프(loop)를 시작합니다.
                Call CATIA_BASIC02()
                If Not FName = "" Then
                    Result_File_Open = CATIA.Documents.Open(Data_Search_Directory & FName)
                    Result_Data_Value(Result_Check_Count) = FName
                    Call Put_Parameter_Data_Comparison_Oil_Piping(Selection_Parameter, Initial_Parameter, Selection_Parameter_Count, Form_Value, Data_Search_Directory)
                End If
                FileListCounter = FileListCounter + 1
                FName = Dir()    ' 다음 항목을 읽어들입니다.

                Result_File_Open.Close()
            Loop
            '--------------------------------------------------------------------------------------------------- 일치하는 Data가 없을때...
            If Result_Check_Count = 1 Then
                MsgBox("일치하는 Data가 없습니다.")
                Form_Value.Message_Label.Caption = "Data 선정이 완료 되었습니다."
            Else
                Dim temp01
                ReDim Result_Data_Value_Number(0 To Form_Value.Data_Path_List.ListCount, 0 To Form_Value.MSFlexGrid1.Cols - 1)
                ReDim temp01(0 To Form_Value.MSFlexGrid1.Cols)
                '--------------------------------------------------------------------------------------------------- 내림차순 정리
                For i = 1 To Form_Value.Data_Path_List.ListCount
                    For j = 1 To Form_Value.MSFlexGrid1.Cols - 1
                        Result_Data_Value_Number(i, j) = Form_Value.MSFlexGrid1.TextMatrix(i, j + 1)
                    Next
                Next
                If Form_Value.Data_Path_List.ListCount > 1 Then
                    '--------------------------------------------------------------------------------------------------- 내림차순 정리하기
                    For i = 1 To Form_Value.Data_Path_List.ListCount - 2
                        For j = i + 1 To Form_Value.Data_Path_List.ListCount - 1
                            If UCase(Result_Data_Value_Number(i, 1)) > UCase(Result_Data_Value_Number(j, 1)) Then
                                For k = 1 To Form_Value.MSFlexGrid1.Cols - 1
                                    temp01(k) = Result_Data_Value_Number(i, k)
                                    Result_Data_Value_Number(i, k) = Result_Data_Value_Number(j, k)
                                    Result_Data_Value_Number(j, k) = temp01(k)
                                    'Exit For
                                Next
                            End If
                        Next
                    Next
                End If
                For i = 1 To Form_Value.Data_Path_List.ListCount
                    For j = 1 To Form_Value.MSFlexGrid1.Cols - 1
                        Form_Value.MSFlexGrid1.TextMatrix(i, j + 1) = Result_Data_Value_Number(i, j)
                    Next
                Next

                '--------------------------------------------------------------------------------------------------- Data Loading
                Call CATIA_BASIC02()
                Dim arrayOfVariantOfBSTR1(0)
                arrayOfVariantOfBSTR1(0) = Form_Value.Data_Path_List.List(0) & Form_Value.MSFlexGrid1.TextMatrix(1, 2)
                CATIA.ActiveDocument.Product.products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

                '--------------------------------------------------------------------------------------------------- Data Loading 안됐을때...
                Dim Result_Data_Value_Name = "Not_Value"
                Result_Data_Value = CATIA.Documents.Item(Form_Value.MSFlexGrid1.TextMatrix(1, 2))
                Result_Data_Value_Name = CATIA.Documents.Item(Form_Value.MSFlexGrid1.TextMatrix(1, 2)).Name
                If Result_Data_Value_Name = "Not_Value" Then
                    MsgBox("Data Open 실패. 실적 Data 정보를 확인하세요")
                Else
                    CATIA.ActiveDocument.Product.products.Item(CATIA.ActiveDocument.Product.products.Count).Name = Result_Template_Item_Value
                End If
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Data_Parameter_Comparison_Oil_Piping() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region








#Region " --- Get_Parameter_Data_Comparison() : 선정시 비교 대상이 될 파라미터/UI의 값(get_parameter_first_value) 추출 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.09
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Parameter_Value : 검색 대상
    '             Loop_Count : 검색 대상들의 전체 개수
    '             oForm : 실행 된 폼의 정보
    '---------------------------------------------------------------------------------------------------
    Public Function Get_Parameter_Data_Comparison(Parameter_Value, Loop_Count, oForm)
        Try
            get_parameter_first_count = 1

            For i = 1 To Loop_Count - 1
                '--------------------------------------------------------------------------------------------------- 축길이 결정일 때.
                If Not Parameter_Value(i, 8) = "1" Then
                    Call axis_xyz_distance(Parameter_Value(i, 1), Parameter_Value(i, 2))

                    If Parameter_Value(i, 5) = "1" Then
                        get_parameter_first_value(get_parameter_first_count) = Math.Round(axis_axis_measure_x, 1)
                        get_parameter_first_count = get_parameter_first_count + 1
                    End If
                    If Parameter_Value(i, 6) = "1" Then
                        get_parameter_first_value(get_parameter_first_count) = Math.Round(axis_axis_measure_y, 1)
                        get_parameter_first_count = get_parameter_first_count + 1
                    End If
                    If Parameter_Value(i, 7) = "1" Then
                        get_parameter_first_value(get_parameter_first_count) = Math.Round(axis_axis_measure_z, 1)
                        get_parameter_first_count = get_parameter_first_count + 1
                    End If
                Else
                    Dim Select_Form_Controls = Nothing

                    If Parameter_Value(i, 2) = "Design_Rule_Version" Then
                        get_parameter_first_value(get_parameter_first_count) = Data_Version_Infomation(1, 1)
                    ElseIf Left(Parameter_Value(i, 2), 3) = "UI_" Then
                        If oForm.text = "B_Coupling_Create" Then
                            '--------------------------------------------------------------------------------------------------- Coupling
                            If Not InStr(Parameter_Value(i, 4), "DISK") = 0 Then
                                '--------------------------------------------------------------------------------------------------- 타입 선택 -> DISK, GEAR, ALL
                                If (oForm.Check1.Checked = True) And (oForm.Check2.Checked = False) Then
                                    get_parameter_first_value(get_parameter_first_count) = "DISK"
                                ElseIf (oForm.Check1.Checked = False) And (oForm.Check2.Checked = True) Then
                                    get_parameter_first_value(get_parameter_first_count) = "GEAR"
                                ElseIf (oForm.Check1.Checked = True) And (oForm.Check2.Checked = True) Then
                                    get_parameter_first_value(get_parameter_first_count) = "ALL"
                                End If
                            ElseIf Not InStr(Parameter_Value(i, 4), "KOR") = 0 Then
                                '--------------------------------------------------------------------------------------------------- 제조사 선택 -> KOR, CHN
                                If oForm.Check1.Checked = True Then                                         ' DISK가 선택된 경우만..
                                    If B_Coupling_Create.CBX_KOR.Checked = True Then
                                        get_parameter_first_value(get_parameter_first_count) = "KOR"
                                    ElseIf B_Coupling_Create.CBX_CHN.Checked = True Then
                                        get_parameter_first_value(get_parameter_first_count) = "CHN"
                                    Else
                                        get_parameter_first_value(get_parameter_first_count) = ""
                                    End If
                                End If
                            End If
                        ElseIf oForm.text = "C_Coupling_Cover_Create" Then
                            '--------------------------------------------------------------------------------------------------- Coupling Cover - 선정
                            '--------------------------------------------------------------------------------------------------- 타입 선택 -> 2, 3, 4
                            get_parameter_first_value(get_parameter_first_count) = oForm.UI_STAGE_COMBO.Text

                            '--------------------------------------------------------------------------------------------------- VVIP
                            If oForm.VVIP_Check.Checked = True Then
                                get_parameter_first_value(get_parameter_first_count) = "true"
                            Else
                                get_parameter_first_value(get_parameter_first_count) = "false"
                            End If
                        ElseIf oForm.text = "D_Oil_Tank_Create" Then
                            '--------------------------------------------------------------------------------------------------- Oil Tank
                            '--------------------------------------------------------------------------------------------------- 타입 선택 -> KOR, CHN
                            If oForm.Option_KOR.Checked = True Then
                                get_parameter_first_value(get_parameter_first_count) = "KOR"
                            Else
                                get_parameter_first_value(get_parameter_first_count) = "CHN"
                            End If
                        ElseIf oForm.text = "DA_Oil_Tank_Template_Create" Then
                            '--------------------------------------------------------------------------------------------------- Oil Tank
                            '--------------------------------------------------------------------------------------------------- 타입 선택 -> KOR, CHN
                            If oForm.Option_KOR.Checked = True Then
                                get_parameter_first_value(get_parameter_first_count) = "KOR"
                            Else
                                get_parameter_first_value(get_parameter_first_count) = "CHN"
                            End If
                        ElseIf STW_Template_Infomation(Form_Click_Index_Value, 8) = "H_Base_Frame_Create" Then
                            '--------------------------------------------------------------------------------------------------- Base Frame
                            ' 돌아가는지, 폼 컨트롤 요소들 가져오는지, 폼 안의 컨트롤 이름 가져오는지 확인 할 것. 190403 이상민
                            Try
                                Select_Form_Controls = oForm.Controls
                                'For Each tmpInner In Select_Form_Controls
                                '    oForm.Controls.Find(Parameter_Value(i, 2), True)(0).Checked
                                '    If tmpInner.Name = Parameter_Value(i, 2) Then
                                '        If tmpInner.Value = True Then
                                '            get_parameter_first_value(get_parameter_first_count) = "true"
                                '        Else
                                '            get_parameter_first_value(get_parameter_first_count) = "false"
                                '        End If
                                '        Exit For
                                '    End If
                                'Next
                                '191129
                                '---------------------------------------------------------------------------------------------------
                                ' 생 성 일  : -
                                ' 생 성 자  : -
                                ' 수 정 일  : 19.11.29
                                ' 수 정 자  : 허혜원
                                ' 수정사유  : 폼 화면 컨트롤(OptionBox 등) 이름으로 가져오는 부분 수정            
                                ' Parameter : -
                                '---------------------------------------------------------------------------------------------------
                                If oForm.Controls.Find(Parameter_Value(i, 2), True)(0).Name = Parameter_Value(i, 2) Then
                                    If oForm.Controls.Find(Parameter_Value(i, 2), True)(0).Checked = True Then
                                        get_parameter_first_value(get_parameter_first_count) = "true"
                                    Else
                                        get_parameter_first_value(get_parameter_first_count) = "false"
                                    End If
                                End If
                            Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)

                            End Try
                        Else
                            '--------------------------------------------------------------------------------------------------- VVIP Check
                            ' 돌아가는지, 폼 컨트롤 요소들 가져오는지, 폼 안의 컨트롤 이름 가져오는지 확인 할 것. 190403 이상민
                            'Select_Form_Controls = oForm.Controls
                            'For k = 1 To Select_Form_Controls.item.Count
                            '    If Select_Form_Controls.item(k).Name = Parameter_Value(i, 2) Then
                            '        If Select_Form_Controls.item("VVIP_Check").checked = True And Select_Form_Controls.Item(k).checked = True Then
                            '            get_parameter_first_value(get_parameter_first_count) = "true"
                            '        Else
                            '            get_parameter_first_value(get_parameter_first_count) = "false"
                            '        End If

                            '        Exit For
                            '    End If
                            'Next k

                            Try
                                'If oForm.Parameter_Value(i, 2).Checked = True Then
                                '    get_parameter_first_value(get_parameter_first_count) = "true"
                                'Else
                                '    get_parameter_first_value(get_parameter_first_count) = "false"
                                'End If
                                '---------------------------------------------------------------------------------------------------
                                ' 생 성 일  : -
                                ' 생 성 자  : -
                                ' 수 정 일  : 19.11.29
                                ' 수 정 자  : 허혜원
                                ' 수정사유  : 폼 화면 컨트롤(OptionBox 등) 이름으로 가져오는 부분 수정            
                                ' Parameter : -
                                '---------------------------------------------------------------------------------------------------
                                If oForm.Controls.Find(Parameter_Value(i, 2), True)(0).Name = Parameter_Value(i, 2) Then
                                    If oForm.Controls.Find(Parameter_Value(i, 2), True)(0).Checked = True Then
                                        get_parameter_first_value(get_parameter_first_count) = "true"
                                    Else
                                        get_parameter_first_value(get_parameter_first_count) = "false"
                                    End If
                                End If
                            Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)

                            End Try
                        End If
                    Else
                        selection.clear
                        selection.search("Name='" & Parameter_Value(i, 2) & "',all")

                        If Not selection.Count = 0 Then
                            get_parameter_first_value(get_parameter_first_count) = selection.Item(1).Value.Value
                        End If
                    End If
                    get_parameter_first_count = get_parameter_first_count + 1
                End If
            Next

            Return True
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Get_Parameter_Data_Comparison 함수 오류!!", "", 1, 1)
            Return False
        End Try
    End Function
#End Region


#Region " --- axis_xyz_distance() : 축(플랜)들 같 거리 측정 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.02
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : axis01_axisname : 거리 측정 대상 1
    '             axis02_axisname : 거리 측정 대상 2
    '---------------------------------------------------------------------------------------------------
    Public Sub axis_xyz_distance(axis01_axisname, axis02_axisname)
        Try
            Dim axis01_center_point(2), axis02_center_point(2)
            Dim axis01_value(11), axis02_value(11)
            Dim x_axis_coord(2), y_axis_coord(2), z_axis_coord(2)
            Dim x_axis_coord2(2), y_axis_coord2(2), z_axis_coord2(2)

            Call CATIA_BASIC02()

            '--------------------------------------------------------------------------------------------------- 새로운 Part 생성
            publications1 = product_add.publications
            constraints1 = product_add.Connections("CATIAConstraints")
            selection = CATIA.ActiveDocument.selection
            selection.clear
            '--------------------------------------------------------------------------------------------------- axis01 value 가져오기
            selection.search("Name='" & axis01_axisname & "',all")
            If selection.count = 0 Then
                MsgBox(axis01_axisname & " 이 존재하지 않습니다.", vbSystemModal)
                Exit Sub
            End If
            Dim axis01_distance_value_selection = selection.Item(1)
            Dim axis01_axissystem As AxisSystem = selection.Item(1).value
            '---------------------------------------------------------------------------------------------------------------- Part Name 찾기
            selection.search("Name='" & selection.Item(1).LeafProduct.PartNumber & "',all")
            Dim axis01_part = selection.Item(1).Value

            Dim i As Integer
            For i = 1 To product_add.products.Count
                If product_add.products.Item(i).PartNumber = axis01_distance_value_selection.LeafProduct.PartNumber Then
                    Exit For
                End If
            Next

            '--------------------------------------------------------------------------------------------------- 기존에 남아있는 불필요한 axis_xyz_distance_01 geometrical set 삭제
            selection.clear
            selection.search("Name='axis_xyz_distance_01*',all")
            If selection.Count > 0 Then
                selection.Delete()
            End If
            axis01_distance_value_selection.Value.iscurrent = True
            axis01_axissystem.GetOrigin(axis01_center_point)

            '--------------------------------------------------------------------------------------------------- 새로운 Geom.set 생성
            Dim geom_set_selection = axis01_part.HybridBodies.Add
            geom_set_selection.Name = "axis_xyz_distance_01"
            '---------------------------------------------------------------------------------------------------------------- axis01_center_point_create 생성
            Dim axis01_center_point_create As HybridShapePointCoord
            If iTemplateFlag = 0 Then
                '---------------------------------------------------------------------------------------------------------------- 실적 검색
                axis01_center_point_create = axis01_part.hybridShapeFactory.AddNewPointCoord(0, 0, 0)
            Else
                '---------------------------------------------------------------------------------------------------------------- 템플릿 생성
                axis01_center_point_create = axis01_part.hybridShapeFactory.AddNewPointCoord(axis01_center_point(0), axis01_center_point(1), axis01_center_point(2))
            End If


            geom_set_selection.AppendHybridShape(axis01_center_point_create)

            Dim Axis_reference = axis01_part.CreateReferenceFromObject(axis01_axissystem)
            axis01_center_point_create.RefAxisSystem = Axis_reference
            axis01_center_point_create.Name = "axis01_center_point"
            '--------------------------------------------------------------------------------------------------- Axis_Value 얻어오기
            axis01_axissystem = axis01_distance_value_selection.Value
            axis01_axissystem.GetXAxis(x_axis_coord)
            axis01_axissystem.GetYAxis(y_axis_coord)
            axis01_axissystem.GetZAxis(z_axis_coord)

            '--------------------------------------------------------------------------------------------------- Create_X_Line
            Dim axis01_x_direction, axis01_x_line
            axis01_x_direction = axis01_part.hybridShapeFactory.AddNewDirectionByCoord(x_axis_coord(0), x_axis_coord(1), x_axis_coord(2))
            axis01_x_line = axis01_part.hybridShapeFactory.AddNewLinePtDir(axis01_center_point_create, axis01_x_direction, 0, 100, False)
            geom_set_selection.AppendHybridShape(axis01_x_line)
            '--------------------------------------------------------------------------------------------------- Create_Y_Line
            Dim axis01_y_direction, axis01_y_line
            axis01_y_direction = axis01_part.hybridShapeFactory.AddNewDirectionByCoord(y_axis_coord(0), y_axis_coord(1), y_axis_coord(2))
            axis01_y_line = axis01_part.hybridShapeFactory.AddNewLinePtDir(axis01_center_point_create, axis01_y_direction, 0, 100, False)
            geom_set_selection.AppendHybridShape(axis01_y_line)
            '--------------------------------------------------------------------------------------------------- Create_Z_Line
            Dim axis01_z_direction, axis01_z_line
            axis01_z_direction = axis01_part.hybridShapeFactory.AddNewDirectionByCoord(z_axis_coord(0), z_axis_coord(1), z_axis_coord(2))
            axis01_z_line = axis01_part.hybridShapeFactory.AddNewLinePtDir(axis01_center_point_create, axis01_z_direction, 0, 100, False)
            geom_set_selection.AppendHybridShape(axis01_z_line)
            '--------------------------------------------------------------------------------------------------- Create_XZ_Plane
            Dim axis01_xz_plane = axis01_part.hybridShapeFactory.AddNewPlane2Lines(axis01_x_line, axis01_z_line)
            geom_set_selection.AppendHybridShape(axis01_xz_plane)
            axis01_xz_plane.Name = "axis01_xz_plane"
            '--------------------------------------------------------------------------------------------------- Create_XY_Plane
            Dim axis01_xy_plane = axis01_part.hybridShapeFactory.AddNewPlane2Lines(axis01_x_line, axis01_y_line)
            geom_set_selection.AppendHybridShape(axis01_xy_plane)
            axis01_xy_plane.Name = "axis01_xy_plane"
            '--------------------------------------------------------------------------------------------------- Create_YZ_Plane
            Dim axis01_yz_plane = axis01_part.hybridShapeFactory.AddNewPlane2Lines(axis01_y_line, axis01_z_line)
            geom_set_selection.AppendHybridShape(axis01_yz_plane)
            axis01_yz_plane.Name = "axis01_yz_plane"
            axis01_part.Update()
            '---------------------------------------------------------------------------------------------------------------- axis01_xz_plane_publication 생성
            Dim reference1, publication1, publication2, publication3
            reference1 = product_add.CreateReferenceFromName(product_add.Name & "/" & product_add.products.Item(i).Name & "/!axis01_xz_plane")
            publication1 = publications1.Add("axis01_xz_plane")
            publications1.SetDirect("axis01_xz_plane", reference1)
            '---------------------------------------------------------------------------------------------------------------- axis01_xy_plane_publication 생성
            reference1 = product_add.CreateReferenceFromName(product_add.Name & "/" & product_add.products.Item(i).Name & "/!axis01_xy_plane")
            publication2 = publications1.Add("axis01_xy_plane")
            publications1.SetDirect("axis01_xy_plane", reference1)
            '---------------------------------------------------------------------------------------------------------------- axis01_xy_plane_publication 생성
            reference1 = product_add.CreateReferenceFromName(product_add.Name & "/" & product_add.products.Item(i).Name & "/!axis01_yz_plane")
            publication3 = publications1.Add("axis01_yz_plane")
            publications1.SetDirect("axis01_yz_plane", reference1)
            selection.clear
            '--------------------------------------------------------------------------------------------------- axis02 value 가져오기
            selection.search("Name='" & axis02_axisname & "',all")
            If selection.count = 0 Then
                MsgBox(axis02_axisname & " 이 존재하지 않습니다.", vbSystemModal)
                Exit Sub
            End If
            '191129
            Dim axis02_distance_value_selection = selection.Item(1)
            Dim axis02_axissystem As AxisSystem = selection.Item(1).value
            '---------------------------------------------------------------------------------------------------------------- Part Name 찾기
            selection.search("Name='" & selection.Item(1).LeafProduct.PartNumber & "',all")
            Dim axis02_part = selection.Item(1).Value
            selection.Clear()
            For i = 1 To product_add.products.Count
                If product_add.products.Item(i).PartNumber = axis02_distance_value_selection.LeafProduct.PartNumber Then
                    Exit For
                End If
            Next

            axis02_distance_value_selection.Value.iscurrent = True
            axis02_axissystem.GetOrigin(axis02_center_point)
            '--------------------------------------------------------------------------------------------------- 새로운 Geom.set 생성
            Dim geom_set_selection02 = axis02_part.HybridBodies.Add
            geom_set_selection02.Name = "axis_xyz_distance_02"
            '---------------------------------------------------------------------------------------------------------------- axis02_center_point_create 생성
            Dim axis02_center_point_create As HybridShapePointCoord
            If iTemplateFlag = 0 Then
                axis02_center_point_create = axis02_part.hybridShapeFactory.AddNewPointCoord(0, 0, 0)
            Else
                axis02_center_point_create = axis02_part.hybridShapeFactory.AddNewPointCoord(axis02_center_point(0), axis02_center_point(1), axis02_center_point(2))
            End If

            geom_set_selection02.AppendHybridShape(axis02_center_point_create)
            Axis_reference = axis02_part.CreateReferenceFromObject(axis02_axissystem)
            axis02_center_point_create.RefAxisSystem = Axis_reference
            axis02_center_point_create.Name = "axis02_center_point"
            '--------------------------------------------------------------------------------------------------- Axis_Value 얻어오기
            axis02_axissystem = axis02_distance_value_selection.Value
            axis02_axissystem.GetXAxis(x_axis_coord2)
            axis02_axissystem.GetYAxis(y_axis_coord2)
            axis02_axissystem.GetZAxis(z_axis_coord2)
            '--------------------------------------------------------------------------------------------------- Create_X_Line
            Dim axis02_x_direction = axis02_part.hybridShapeFactory.AddNewDirectionByCoord(x_axis_coord2(0), x_axis_coord2(1), x_axis_coord2(2))
            Dim axis02_x_line = axis02_part.hybridShapeFactory.AddNewLinePtDir(axis02_center_point_create, axis02_x_direction, 0, 100, False)
            geom_set_selection02.AppendHybridShape(axis02_x_line)
            '--------------------------------------------------------------------------------------------------- Create_Y_Line
            Dim axis02_y_direction = axis02_part.hybridShapeFactory.AddNewDirectionByCoord(y_axis_coord2(0), y_axis_coord2(1), y_axis_coord2(2))
            Dim axis02_y_line = axis02_part.hybridShapeFactory.AddNewLinePtDir(axis02_center_point_create, axis02_y_direction, 0, 100, False)
            geom_set_selection02.AppendHybridShape(axis02_y_line)
            '--------------------------------------------------------------------------------------------------- Create_Z_Line
            Dim axis02_z_direction = axis02_part.hybridShapeFactory.AddNewDirectionByCoord(z_axis_coord2(0), z_axis_coord2(1), z_axis_coord2(2))
            Dim axis02_z_line = axis02_part.hybridShapeFactory.AddNewLinePtDir(axis02_center_point_create, axis02_z_direction, 0, 100, False)
            geom_set_selection02.AppendHybridShape(axis02_z_line)
            '--------------------------------------------------------------------------------------------------- Create_XZ_Plane
            Dim axis02_xz_plane = axis02_part.hybridShapeFactory.AddNewPlane2Lines(axis02_x_line, axis02_z_line)
            geom_set_selection02.AppendHybridShape(axis02_xz_plane)
            axis02_xz_plane.Name = "axis02_xz_plane"
            '--------------------------------------------------------------------------------------------------- Create_XY_Plane
            Dim axis02_xy_plane = axis02_part.hybridShapeFactory.AddNewPlane2Lines(axis02_x_line, axis02_y_line)
            geom_set_selection02.AppendHybridShape(axis02_xy_plane)
            axis02_xy_plane.Name = "axis02_xy_plane"
            '--------------------------------------------------------------------------------------------------- Create_YZ_Plane
            Dim axis02_yz_plane = axis02_part.hybridShapeFactory.AddNewPlane2Lines(axis02_y_line, axis02_z_line)
            geom_set_selection02.AppendHybridShape(axis02_yz_plane)
            axis02_yz_plane.Name = "axis02_yz_plane"
            axis02_part.Update()
            '---------------------------------------------------------------------------------------------------------------- axis01_xz_plane_publication 생성
            reference1 = product_add.CreateReferenceFromName(product_add.Name & "/" & product_add.products.Item(i).Name & "/!axis02_xz_plane")
            Dim publication4 = publications1.Add("axis02_xz_plane")
            publications1.SetDirect("axis02_xz_plane", reference1)
            '---------------------------------------------------------------------------------------------------------------- axis01_xy_plane_publication 생성
            reference1 = product_add.CreateReferenceFromName(product_add.Name & "/" & product_add.products.Item(i).Name & "/!axis02_yz_plane")
            Dim publication5 = publications1.Add("axis02_yz_plane")
            publications1.SetDirect("axis02_yz_plane", reference1)
            '---------------------------------------------------------------------------------------------------------------- axis01_xy_plane_publication 생성
            reference1 = product_add.CreateReferenceFromName(product_add.Name & "/" & product_add.products.Item(i).Name & "/!axis02_xy_plane")
            Dim publication6 = publications1.Add("axis02_xy_plane")
            publications1.SetDirect("axis02_xy_plane", reference1)
            product_add.Update()

            '---------------------------------------------------------------------------------------------------  Axis Y Value 가져오기
            Pub1_DisplayName = publication1.Valuation.DisplayName
            Pub2_DisplayName = publication4.Valuation.DisplayName
            Pub_reference1 = product_add.CreateReferenceFromName(Pub1_DisplayName)
            Pub_reference2 = product_add.CreateReferenceFromName(Pub2_DisplayName)
            '--------------------------------------------------------------------------------------------------- Length Measure pyo
            Dim TheSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
            Dim TheMeasurable = TheSPAWorkbench.GetMeasurable(Pub_reference1)
            axis_axis_measure_y = Math.Round(TheMeasurable.GetMinimumDistance(Pub_reference2), 1)
            '---------------------------------------------------------------------------------------------------  Axis Z Value 가져오기
            Pub1_DisplayName = publication2.Valuation.DisplayName
            Pub2_DisplayName = publication6.Valuation.DisplayName
            Pub_reference1 = product_add.CreateReferenceFromName(Pub1_DisplayName)
            Pub_reference2 = product_add.CreateReferenceFromName(Pub2_DisplayName)
            '--------------------------------------------------------------------------------------------------- Length Measure
            TheSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
            TheMeasurable = TheSPAWorkbench.GetMeasurable(Pub_reference1)
            axis_axis_measure_z = Math.Round(TheMeasurable.GetMinimumDistance(Pub_reference2), 1)
            '---------------------------------------------------------------------------------------------------  Axis X Value 가져오기
            Pub1_DisplayName = publication3.Valuation.DisplayName
            Pub2_DisplayName = publication5.Valuation.DisplayName
            Pub_reference1 = product_add.CreateReferenceFromName(Pub1_DisplayName)
            Pub_reference2 = product_add.CreateReferenceFromName(Pub2_DisplayName)
            '--------------------------------------------------------------------------------------------------- Length Measure
            TheSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
            TheMeasurable = TheSPAWorkbench.GetMeasurable(Pub_reference1)
            axis_axis_measure_x = Math.Round(TheMeasurable.GetMinimumDistance(Pub_reference2), 1)
            '--------------------------------------------------------------------------------------------------- Data Delete
            selection.Clear()
            publications1.Remove("axis01_xz_plane")
            publications1.Remove("axis01_xy_plane")
            publications1.Remove("axis01_yz_plane")
            publications1.Remove("axis02_xz_plane")
            publications1.Remove("axis02_xy_plane")
            publications1.Remove("axis02_yz_plane")
            selection.Add(geom_set_selection)
            selection.Add(geom_set_selection02)
            selection.Delete()
            product_add.Update()
        Catch ex As Exception : MessageBox.Show(ex.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("axis_xyz_distance 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Put_Parameter_Data_Comparison() : 선정시 파라미터 값들을 비교하여 선정대상이 맞는지 확인 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.02
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Parameter_Value01 : 기종_Constraint_List.xlsx 에서 읽은 정보
    '             Parameter_Value02 : 
    '             Loop_Count : 비교 할 파라미터 개수
    '             VVIP_Form : 실행 한 폼 정보
    '             Data_Path : 열려있는 템플릿의 경로
    '---------------------------------------------------------------------------------------------------
    Public Sub Put_Parameter_Data_Comparison(Parameter_Value01, Parameter_Value02, Loop_Count, VVIP_Form, Data_Path)
        Try
            Call CATIA_BASIC02()

            get_parameter_second_count = 1
            data_open_count = 1
            Dim j = 1
            Dim xyz_number

            For i = 1 To Loop_Count - 1
                If i > 1 Then
                    For j = 1 To Loop_Count - 1
                        If Parameter_Value01(i, 1) = Parameter_Value02(j, 3) Then
                            For k = 1 To j
                                '--------------------------------------------------------------------------------------------------- x값이 1일때
                                If Parameter_Value01(k, 5) = 1 Then
                                    '--------------------------------------------------------------------------------------------------- x,y 값이 1일때
                                    If Parameter_Value01(k, 6) = 1 Then
                                        j = j + 1
                                        '--------------------------------------------------------------------------------------------------- x,y,z 값이 모두 1일때
                                        If Parameter_Value01(k, 7) = 1 Then
                                            j = j + 1
                                        End If
                                        '--------------------------------------------------------------------------------------------------- x,z 값이 1일때
                                    ElseIf Parameter_Value01(k, 7) = 1 Then
                                        j = j + 1
                                    End If
                                End If
                                '--------------------------------------------------------------------------------------------------- y,z 값이 1일때
                                If Parameter_Value01(k, 6) = 1 Then
                                    If Parameter_Value01(k, 7) = 1 Then
                                        j = j + 1
                                    End If
                                End If
                            Next
                            Exit For
                        End If
                    Next
                End If

                '--------------------------------------------------------------------------------------------------- XYZ 측정 체크
                xyz_number = 0
                If Parameter_Value01(i, 5) = "1" Then
                    xyz_number = xyz_number + 1
                End If
                If Parameter_Value01(i, 6) = "1" Then
                    xyz_number = xyz_number + 1
                End If
                If Parameter_Value01(i, 7) = "1" Then
                    xyz_number = xyz_number + 1
                End If

                '--------------------------------------------------------------------------------------------------- 축 길이 측정
                If Not Parameter_Value01(i, 8) = "1" Then
                    '--------------------------------------------------------------------------------------------------- X 축 길이 측정
                    If Parameter_Value01(i, 5) = "1" Then
                        If xyz_number = 1 Then
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                        Else
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "_X',all")
                        End If
                        '--------------------------------------------------------------------------------------------------- Parameter가 없을때
                        If selection.Count = 0 Then
                            data_open_count = 0
                            Exit For
                        End If

                        get_parameter_second_value(get_parameter_second_count) = Math.Round(selection.Item(1).Value.Value, 1)
                        '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                        If Parameter_Value01(i, 9) <= -1 Then
                            If Not get_parameter_first_value(j) >= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                        ElseIf Parameter_Value01(i, 9) = 0 Then
                            If Not get_parameter_first_value(j) = Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                        ElseIf Parameter_Value01(i, 9) >= 1 Then
                            If Not get_parameter_first_value(j) <= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                        End If
                        get_parameter_second_count = get_parameter_second_count + 1
                        j = j + 1
                    End If
                    '--------------------------------------------------------------------------------------------------- Y 축 길이 측정
                    If Parameter_Value01(i, 6) = "1" Then
                        If xyz_number = 1 Then
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                        Else
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "_Y',all")
                        End If
                        '--------------------------------------------------------------------------------------------------- Parameter가 없을때
                        If selection.Count = 0 Then
                            data_open_count = 0
                            Exit For
                        End If
                        get_parameter_second_value(get_parameter_second_count) = Math.Round(selection.Item(1).Value.Value, 1)
                        '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                        If Parameter_Value01(i, 9) <= -1 Then
                            If Not get_parameter_first_value(j) >= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                        ElseIf Parameter_Value01(i, 9) = 0 Then
                            If Not get_parameter_first_value(j) = Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                        ElseIf Parameter_Value01(i, 9) >= 1 Then
                            If Not get_parameter_first_value(j) <= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                        End If
                        get_parameter_second_count = get_parameter_second_count + 1
                        j = j + 1
                    End If
                    '--------------------------------------------------------------------------------------------------- Z 축 길이 측정
                    If Parameter_Value01(i, 7) = "1" Then
                        If xyz_number = 1 Then
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                        Else
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "_Z',all")
                        End If
                        '--------------------------------------------------------------------------------------------------- Parameter가 없을때
                        If selection.Count = 0 Then
                            data_open_count = 0
                            Exit For
                        End If
                        get_parameter_second_value(get_parameter_second_count) = Math.Round(selection.Item(1).Value.Value, 1)
                        '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                        If Parameter_Value01(i, 9) <= -1 Then
                            If Not get_parameter_first_value(j) >= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                        ElseIf Parameter_Value01(i, 9) = 0 Then
                            If Not get_parameter_first_value(j) = Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                        ElseIf Parameter_Value01(i, 9) >= 1 Then
                            If Not get_parameter_first_value(j) <= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                        End If
                        get_parameter_second_count = get_parameter_second_count + 1
                    End If
                Else
                    If Parameter_Value01(i, 10) = "STD" Then
                        selection.clear
                        selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                        '--------------------------------------------------------------------------------------------------- Parameter가 없을때
                        If selection.Count = 0 Then
                            data_open_count = 0
                            Exit For
                        End If

                        get_parameter_second_value(get_parameter_second_count) = selection.Item(1).Value.Value

                        '--------------------------------------------------------------------------------------------------- 파라미터 값이 더블이고, 소수점 자리수가 길게 나오는경우는 반올림 처리
                        Dim bItemType = get_parameter_second_value(get_parameter_second_count).GetType
                        If bItemType.Name = "Double" Then
                            get_parameter_second_value(get_parameter_second_count) = Math.Round(get_parameter_second_value(get_parameter_second_count), 1)
                        End If

                        If Parameter_Value01(i, 1) = "UI_CONTROL_PANEL_TYPE" Then

                            Dim asda = ""
                            ' UI_CONTROL_PANEL_TYPE 폼 생성 후 작업 해야함

                            'If L_Motor_Support_Create.Motor_Support_CTRL_PNL_OP1.Value = True Then
                            '    If Not get_parameter_second_value(get_parameter_second_count) = "700X1350X300" Then
                            '        data_open_count = 0
                            '        Exit For
                            '    End If
                            'Else
                            '    If Not get_parameter_second_value(get_parameter_second_count) = "840X1850X340" Then
                            '        data_open_count = 0
                            '        Exit For
                            '    End If
                            'End If
                        Else
                            '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                            If Parameter_Value01(i, 9) <= -1 Then
                                If Not get_parameter_first_value(j) >= get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                                '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                            ElseIf Parameter_Value01(i, 9) = 0 Then
                                '--------------------------------------------------------------------------------------------------- Layout에서 값이 없을때... 값을 false로 변경(140822)
                                '--------------------------------------------------------------------------------------------------- 구분값이 "KOR" or "CHN"
                                If CStr(get_parameter_second_value(get_parameter_second_count)) = "KOR" Or CStr(get_parameter_second_value(get_parameter_second_count)) = "CHN" Then
                                    If D_Oil_Tank_Create.Option_KOR.Checked = True Then
                                        get_parameter_first_value(j) = "KOR"
                                    ElseIf D_Oil_Tank_Create.Option_KOR.Checked = True Then
                                        get_parameter_first_value(j) = "CHN"
                                    End If
                                End If
                                '--------------------------------------------------------------------------------------------------- 값이 없으면 "X"
                                If CStr(get_parameter_first_value(j)) = "" Then
                                    get_parameter_first_value(j) = "X"
                                End If
                                '--------------------------------------------------------------------------------------------------- get_parameter_first_value = get_parameter_second_value ?
                                If Not get_parameter_first_value(j) = get_parameter_second_value(get_parameter_second_count) Then
                                    If CStr(get_parameter_first_value(j)) = "ALL" And CStr(get_parameter_second_value(get_parameter_second_count)) = "DISK" Or
                                       CStr(get_parameter_second_value(get_parameter_second_count)) = "GEAR" Then
                                    Else
                                        data_open_count = 0
                                        Exit For
                                    End If
                                End If
                                '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                            ElseIf Parameter_Value01(i, 9) >= 1 Then
                                If Not get_parameter_first_value(j) <= get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                            End If
                        End If
                    Else
                        '---------------------------------------------------------------------------------------------------
                        ' VVIP 일때....
                        '---------------------------------------------------------------------------------------------------
                        Dim VVIP_Check_Value, VVIP_Check_Item_Value

                        If VVIP_Form.VVIP_Check.Checked = False Then
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                            VVIP_Check_Value = "false"
                            VVIP_Check_Value = selection.Item(1).Value.Value
                            get_parameter_second_value(get_parameter_second_count) = selection.Item(1).Value.Value
                            If VVIP_Check_Value = "true" Or VVIP_Check_Value = "True" Then
                                data_open_count = 0
                            ElseIf VVIP_Check_Value = "false" Or VVIP_Check_Value = "False" Then
                                get_parameter_second_value(get_parameter_second_count) = "false"
                            End If
                        ElseIf VVIP_Form.VVIP_Check.Checked = True Then
                            VVIP_Check_Item_Value = C_Coupling_Cover_Create.Controls.Item(Parameter_Value01(i, 1))
                            If VVIP_Check_Item_Value.Checked = True Then
                                get_parameter_first_value(j) = "true"
                            Else
                                get_parameter_first_value(j) = "false"
                            End If
                            get_parameter_second_value(get_parameter_second_count) = "false"
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")

                            If selection.Item(1).Value.Value = "True" Or selection.Item(1).Value.Value = "true" Then
                                get_parameter_second_value(get_parameter_second_count) = "true"
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                            If Parameter_Value01(i, 9) <= -1 Then
                                If Not get_parameter_first_value(j) >= get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                                '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                            ElseIf Parameter_Value01(i, 9) = 0 Then
                                '--------------------------------------------------------------------------------------------------- Layout에서 값이 없을때... 값을 false로 변경
                                If get_parameter_first_value(j) = Nothing Or get_parameter_first_value(j) = "" Then
                                    get_parameter_first_value(j) = "X"
                                End If
                                If Not get_parameter_first_value(j) = get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                                '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                            ElseIf Parameter_Value01(i, 9) >= 1 Then
                                If Not get_parameter_first_value(j) <= get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                    get_parameter_second_count = get_parameter_second_count + 1
                End If
            Next

            '--------------------------------------------------------------------------------------------------- 선정 조건에 일치하면 값 기입
            If data_open_count = 1 Then
                If Result_Check_Count = 1 Then
                    VVIP_Form.DataGridView1.Rows(Result_Check_Count - 1).Cells(0).Value = ">"
                End If

                VVIP_Form.Data_Path_List.Items.Add(Data_Path)

                VVIP_Form.DataGridView1.RowCount = Result_Check_Count
                VVIP_Form.DataGridView1.Rows(Result_Check_Count - 1).Cells(1).Value = Result_Check_Count
                VVIP_Form.DataGridView1.Rows(Result_Check_Count - 1).Cells(2).Value = Result_Data_Value(Result_Check_Count)
                For i = 1 To get_parameter_second_count - 1
                    VVIP_Form.DataGridView1.Rows(Result_Check_Count - 1).Cells(i + 2).Value = get_parameter_second_value(i)
                Next

                VVIP_Form.DataGridView1.Refresh()
                Result_Check_Count = Result_Check_Count + 1
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Put_Parameter_Data_Comparison() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Put_Parameter_Data_Comparison_Oil_Piping() : 선정시 파라미터 값들을 비교하여 선정대상이 맞는지 확인 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.02
    ' 수 정 자  : 이상민
    ' 수정사유  : 함수 개선 및 오류 처리
    ' Parameter : Parameter_Value01 : 기종_Constraint_List.xlsx 에서 읽은 정보
    '             Parameter_Value02 : 
    '             Loop_Count : 비교 할 파라미터 개수
    '             VVIP_Form : 실행 한 폼 정보
    '             Data_Path : 열려있는 템플릿의 경로
    '---------------------------------------------------------------------------------------------------
    Public Sub Put_Parameter_Data_Comparison_Oil_Piping(Parameter_Value01, Parameter_Value02, Loop_Count, VVIP_Form, Data_Path)
        Try
            Call CATIA_BASIC02()
            get_parameter_second_count = 1
            data_open_count = 1
            Dim j = 1

            For i = 1 To Loop_Count - 1
                If i > 1 Then
                    For j = 1 To Loop_Count - 1
                        If Parameter_Value01(i, 1) = Parameter_Value02(j, 3) Then
                            For k = 1 To j
                                '--------------------------------------------------------------------------------------------------- x값이 1일때
                                If Parameter_Value01(k, 5) = 1 Then
                                    '--------------------------------------------------------------------------------------------------- x,y 값이 1일때
                                    If Parameter_Value01(k, 6) = 1 Then
                                        j = j + 1
                                        '--------------------------------------------------------------------------------------------------- x,y,z 값이 모두 1일때
                                        If Parameter_Value01(k, 7) = 1 Then
                                            j = j + 1
                                        End If
                                        '--------------------------------------------------------------------------------------------------- x,z 값이 1일때
                                    ElseIf Parameter_Value01(k, 7) = 1 Then
                                        j = j + 1
                                    End If
                                End If
                                '--------------------------------------------------------------------------------------------------- y,z 값이 1일때
                                If Parameter_Value01(k, 6) = 1 Then
                                    If Parameter_Value01(k, 7) = 1 Then
                                        j = j + 1
                                    End If
                                End If
                            Next
                            Exit For
                        End If
                    Next
                End If
                '--------------------------------------------------------------------------------------------------- XYZ 측정 체크
                Dim xyz_number = 0
                If Parameter_Value01(i, 5) = "1" Then
                    xyz_number = xyz_number + 1
                End If
                If Parameter_Value01(i, 6) = "1" Then
                    xyz_number = xyz_number + 1
                End If
                If Parameter_Value01(i, 7) = "1" Then
                    xyz_number = xyz_number + 1
                End If
                '--------------------------------------------------------------------------------------------------- 축 길이 측정
                If Not Parameter_Value01(i, 8) = "1" Then
                    '--------------------------------------------------------------------------------------------------- X 축 길이 측정
                    If Parameter_Value01(i, 5) = "1" Then
                        If xyz_number = 1 Then
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                        Else
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "_X',all")
                        End If
                        '--------------------------------------------------------------------------------------------------- Parameter가 없을때
                        If selection.Count = 0 Then
                            data_open_count = 0
                            Exit For
                        End If

                        get_parameter_second_value(get_parameter_second_count) = Math.Round(selection.Item(1).Value.Value, 1)
                        '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                        If Parameter_Value01(i, 9) <= -1 Then
                            If Not get_parameter_first_value(j) >= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                        ElseIf Parameter_Value01(i, 9) = 0 Then
                            If i = 1 Then
                                Dim Range_Value = Nothing


                                ' 돌아가는지, 불필요한부분 제거 필요 / lsm / 190520

                                ' 주석 풀어서 확인해야 함
                                'On Error Resume Next
                                'Err.Clear()
                                'Range_Value = CInt(E_Oil_Piping_Create.Range.Text)
                                'If Err.Number <> 0 Then
                                '    data_open_count = 0
                                '    Exit For
                                'End If
                                'On Error GoTo 0



                                '---------------------------------------------------------------------------------------------------

                                ' 돌아가는지, 불필요한부분 제거 필요 / lsm / 190520


                                '   생 성 일 : 18-04-19
                                '   생 성 자 : 허혜원
                                '   수 정 일 :
                                '   수 정 자 :
                                '   수정사유 :
                                '   기    능 : Oil Piping 실적 검색시 Range 값 +,- 구분하여 범위 내의 Oil Piping 을 모두 검색한다.
                                '---------------------------------------------------------------------------------------------------
                                If Range_Value > 0 Then         ' +일때
                                    If (Not get_parameter_first_value(j) <= Math.Round(get_parameter_second_value(get_parameter_second_count), 1)) Or (Not get_parameter_first_value(j) + Range_Value >= Math.Round(get_parameter_second_value(get_parameter_second_count), 1)) Then
                                        data_open_count = 0
                                        Exit For
                                    End If
                                ElseIf Range_Value < 0 Then     ' -일때
                                    If (Not get_parameter_first_value(j) >= Math.Round(get_parameter_second_value(get_parameter_second_count), 1)) Or (Not get_parameter_first_value(j) + Range_Value <= Math.Round(get_parameter_second_value(get_parameter_second_count), 1)) Then
                                        data_open_count = 0
                                        Exit For
                                    End If
                                End If
                            Else
                                If Not get_parameter_first_value(j) = Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                        ElseIf Parameter_Value01(i, 9) >= 1 Then
                            If Not get_parameter_first_value(j) <= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                        End If
                        get_parameter_second_count = get_parameter_second_count + 1
                        j = j + 1
                    End If
                    '--------------------------------------------------------------------------------------------------- Y 축 길이 측정
                    If Parameter_Value01(i, 6) = "1" Then
                        If xyz_number = 1 Then
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                        Else
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "_Y',all")
                        End If
                        '--------------------------------------------------------------------------------------------------- Parameter가 없을때
                        If selection.Count = 0 Then
                            data_open_count = 0
                            Exit For
                        End If
                        get_parameter_second_value(get_parameter_second_count) = Math.Round(selection.Item(1).Value.Value, 1)
                        '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                        If Parameter_Value01(i, 9) <= -1 Then
                            If Not get_parameter_first_value(j) >= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                        ElseIf Parameter_Value01(i, 9) = 0 Then
                            If Not get_parameter_first_value(j) = Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                        ElseIf Parameter_Value01(i, 9) >= 1 Then
                            If Not get_parameter_first_value(j) <= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                        End If
                        get_parameter_second_count = get_parameter_second_count + 1
                        j = j + 1
                    End If
                    '--------------------------------------------------------------------------------------------------- Z 축 길이 측정
                    If Parameter_Value01(i, 7) = "1" Then
                        If xyz_number = 1 Then
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                        Else
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "_Z',all")
                        End If
                        '--------------------------------------------------------------------------------------------------- Parameter가 없을때
                        If selection.Count = 0 Then
                            data_open_count = 0
                            Exit For
                        End If
                        get_parameter_second_value(get_parameter_second_count) = Math.Round(selection.Item(1).Value.Value, 1)
                        '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                        If Parameter_Value01(i, 9) <= -1 Then
                            If Not get_parameter_first_value(j) >= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                        ElseIf Parameter_Value01(i, 9) = 0 Then
                            If Not get_parameter_first_value(j) = Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                        ElseIf Parameter_Value01(i, 9) >= 1 Then
                            If Not get_parameter_first_value(j) <= Math.Round(get_parameter_second_value(get_parameter_second_count), 1) Then
                                data_open_count = 0
                                Exit For
                            End If
                        End If
                        get_parameter_second_count = get_parameter_second_count + 1
                    End If
                Else
                    If Parameter_Value01(i, 10) = "STD" Then
                        selection.clear
                        selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                        '--------------------------------------------------------------------------------------------------- Parameter가 없을때
                        If selection.Count = 0 Then
                            data_open_count = 0
                            Exit For
                        End If

                        get_parameter_second_value(get_parameter_second_count) = selection.Item(1).Value.Value

                        If Parameter_Value01(i, 1) = "UI_CONTROL_PANEL_TYPE" Then

                            ' UI_CONTROL_PANEL_TYPE 폼 생성 후 작업 해야함

                            'If L_Motor_Support_Create.Motor_Support_CTRL_PNL_OP1.Value = True Then
                            '    If Not get_parameter_second_value(get_parameter_second_count) = "700X1350X300" Then
                            '        data_open_count = 0
                            '        Exit For
                            '    End If
                            'Else
                            '    If Not get_parameter_second_value(get_parameter_second_count) = "840X1850X340" Then
                            '        data_open_count = 0
                            '        Exit For
                            '    End If
                            'End If
                        Else
                            '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                            If Parameter_Value01(i, 9) <= -1 Then
                                If Not get_parameter_first_value(j) >= get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                                '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                            ElseIf Parameter_Value01(i, 9) = 0 Then
                                '--------------------------------------------------------------------------------------------------- Layout에서 값이 없을때... 값을 false로 변경(140822)
                                '--------------------------------------------------------------------------------------------------- 구분값이 "KOR" or "CHN"
                                If get_parameter_second_value(get_parameter_second_count) = "KOR" Or get_parameter_second_value(get_parameter_second_count) = "CHN" Then
                                    If D_Oil_Tank_Create.Option_KOR.Checked = True Then
                                        get_parameter_first_value(j) = "KOR"
                                    Else
                                        get_parameter_first_value(j) = "CHN"
                                    End If
                                End If
                                '--------------------------------------------------------------------------------------------------- 값이 없으면 "X"
                                If get_parameter_first_value(j) = Nothing Or get_parameter_first_value(j) = "" Then
                                    get_parameter_first_value(j) = "X"
                                End If
                                '--------------------------------------------------------------------------------------------------- get_parameter_first_value = get_parameter_second_value ?
                                If Not get_parameter_first_value(j) = get_parameter_second_value(get_parameter_second_count) Then
                                    If get_parameter_first_value(j) = "ALL" And (get_parameter_second_value(get_parameter_second_count) = "DISK" Or get_parameter_second_value(get_parameter_second_count) = "GEAR") Then
                                    Else
                                        data_open_count = 0
                                        Exit For
                                    End If
                                End If
                                '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                            ElseIf Parameter_Value01(i, 9) >= 1 Then
                                If Not get_parameter_first_value(j) <= get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                            End If
                        End If
                    Else
                        '---------------------------------------------------------------------------------------------------
                        ' VVIP 일때....
                        '---------------------------------------------------------------------------------------------------
                        Dim VVIP_Check_Value, VVIP_Check_Item_Value

                        If VVIP_Form.VVIP_Check.Checked = False Then
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")
                            VVIP_Check_Value = "false"
                            VVIP_Check_Value = selection.Item(1).Value.Value
                            get_parameter_second_value(get_parameter_second_count) = selection.Item(1).Value.Value
                            If VVIP_Check_Value = "true" Or VVIP_Check_Value = "True" Then
                                data_open_count = 0
                            ElseIf VVIP_Check_Value = "false" Or VVIP_Check_Value = "False" Then
                                get_parameter_second_value(get_parameter_second_count) = "false"
                            End If
                        End If
                        If VVIP_Form.VVIP_Check.Checked = True Then
                            VVIP_Check_Item_Value = C_Coupling_Cover_Create.Controls.Item(Parameter_Value01(i, 1))
                            If VVIP_Check_Item_Value.Checked = True Then
                                get_parameter_first_value(j) = "true"
                            Else
                                get_parameter_first_value(j) = "false"
                            End If
                            get_parameter_second_value(get_parameter_second_count) = "false"
                            selection.clear
                            selection.search("Name='" & Parameter_Value01(i, 2) & "',all")

                            If selection.Item(1).Value.Value = "True" Or selection.Item(1).Value.Value = "true" Then
                                get_parameter_second_value(get_parameter_second_count) = "true"
                            End If
                            '--------------------------------------------------------------------------------------------------- 비교값이 -1 일때
                            If Parameter_Value01(i, 9) <= -1 Then
                                If Not get_parameter_first_value(j) >= get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                                '--------------------------------------------------------------------------------------------------- 비교값이 0 일때
                            ElseIf Parameter_Value01(i, 9) = 0 Then
                                '--------------------------------------------------------------------------------------------------- Layout에서 값이 없을때... 값을 false로 변경
                                If get_parameter_first_value(j) = Nothing Or get_parameter_first_value(j) = "" Then
                                    get_parameter_first_value(j) = "X"
                                End If
                                If Not get_parameter_first_value(j) = get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                                '--------------------------------------------------------------------------------------------------- 비교값이 1 일때
                            ElseIf Parameter_Value01(i, 9) >= 1 Then
                                If Not get_parameter_first_value(j) <= get_parameter_second_value(get_parameter_second_count) Then
                                    data_open_count = 0
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                    get_parameter_second_count = get_parameter_second_count + 1
                End If
            Next

            '--------------------------------------------------------------------------------------------------- 선정 조건에 일치하면 값 기입
            If data_open_count = 1 Then
                If Result_Check_Count = 1 Then
                    VVIP_Form.DataGridView1.Rows(Result_Check_Count - 1).Cells(0).Value = ">"
                End If

                VVIP_Form.Data_Path_List.Items.Add(Data_Path)

                VVIP_Form.DataGridView1.RowCount = Result_Check_Count
                VVIP_Form.DataGridView1.Rows(Result_Check_Count - 1).Cells(1).Value = Result_Check_Count
                VVIP_Form.DataGridView1.Rows(Result_Check_Count - 1).Cells(2).Value = Result_Data_Value(Result_Check_Count)
                For i = 1 To get_parameter_second_count - 1
                    VVIP_Form.DataGridView1.Rows(Result_Check_Count - 1).Cells(i + 2).Value = get_parameter_second_value(i)
                Next

                Result_Check_Count = Result_Check_Count + 1
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Put_Parameter_Data_Comparison_Oil_Piping() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- VVIP 옵션 선택에 따른 폼 크기 조정"
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.08
    ' 수 정 자  : 이상민
    ' 수정사유  : 컨트롤 속성 값 변경에 따른 매개변수 값 변경
    ' Parameter : 
    '---------------------------------------------------------------------------------------------------
    Public Sub VVIP_Check_Form_Size(Min_Size, Max_Size, Form_Value)
        Try
            If Form_Value.VVIP_Check.Checked = True Then
                Form_Value.Width = Max_Size
                Form_Value.VVIP_Frame.Visible = True
            Else
                Form_Value.Width = Min_Size
                Form_Value.VVIP_Frame.Visible = False
            End If
        Catch ex As Exception : MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Dim asda = ""
        End Try
    End Sub

#End Region

End Module
