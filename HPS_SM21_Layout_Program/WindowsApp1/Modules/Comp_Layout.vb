Imports Scripting
Imports MECMOD
Imports HybridShapeTypeLib
Imports DRAFTINGITF


Module Comp_Layout
    '--------------------------------------------------------------------------------------------------- 프로그램 Update 관련 정보
    Public Comp_Layout_Update_Value
    '--------------------------------------------------------------------------------------------------- Excel 수량 Count
    Public processes_count_value
    '--------------------------------------------------------------------------------------------------- 현재 생성되는 Project 기종 정보
    Public assy_value_Name, Label_assy_value_Name
    '--------------------------------------------------------------------------------------------------- 현재 생성되는 Project Sales_No 정보
    Public Sales_No
    '--------------------------------------------------------------------------------------------------- Coupling 유무 판별
    Public coupling_number
    '--------------------------------------------------------------------------------------------------- Oil_Piping_Part Name 정의 ( Oil Piping Layout에 사용 )
    Public Oil_Piping_Part_Value
    '--------------------------------------------------------------------------------------------------- STW_COMP_Infomation 가져오기
    Public STW_COMP_Infomation, STW_COMP_Infomation_Count
    '--------------------------------------------------------------------------------------------------- STW_Template_Infomation_Count 가져오기
    Public STW_Template_Infomation_Count, STW_Template_Infomation
    '--------------------------------------------------------------------------------------------------- Manual 관련 정보 가져오기
    Public Manual_Infomation_Count, Manual_Infomation
    '--------------------------------------------------------------------------------------------------- COMP_Product_File_Name 가져오기
    Public COMP_Product_File_Name
    '--------------------------------------------------------------------------------------------------- 화면 선택 Delay
    Public Click_Count
    '--------------------------------------------------------------------------------------------------- Base 도면 생성
    Public drawingDocument1, drawing_selection, drawingSheets1, drawingSheet1, drawingViews1
    Public Base_Drawing_Front_View, bullgear_value
    '--------------------------------------------------------------------------------------------------- Base 도면 Axis Value
    Public View_Axis_Value(6)
    '--------------------------------------------------------------------------------------------------- 
    Public Drawing_RepresentationMode_Value

    Public Enclosure_Value = Nothing


#Region " --- Find_data_Path() : Form Loading시 Data Path를 자동으로 찾는 기능 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.05.03
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Find_data_Path()
        Try
            Exc = CreateObject("excel.application")
            Exc.Workbooks.Open(Application.StartupPath & "\data_path.xlsx")

            A_Main_Form.Directory_Path_Text.Text = Exc.Worksheets("Sheet1").Range("A1").Value
            A_Main_Form.Constraint_Excel_Path_Text.Text = Exc.Worksheets("Sheet1").Range("A2").Value
            A_Main_Form.O_BOM_Path_Text.Text = Exc.Worksheets("Sheet1").Range("A3").Value
            A_Main_Form.Project_DB_Path_Text.Text = Exc.Worksheets("Sheet1").Range("A4").Value
            A_Main_Form.Template_Data_Path_Text.Text = Exc.Worksheets("Sheet1").Range("A16").Value

            '--------------------------------------------------------------------------------------------------- 메인 폼의 경로들 변수화
            sDirectory_Path_Text = Exc.Worksheets("Sheet1").Range("A1").Value
            sConstraint_Excel_Path_Text = Exc.Worksheets("Sheet1").Range("A2").Value
            sO_BOM_Path_text = Exc.Worksheets("Sheet1").Range("A3").Value
            sProject_DB_Path_text = Exc.Worksheets("Sheet1").Range("A4").Value
            sTemplate_Data_Path_text = Exc.Worksheets("Sheet1").Range("A16").Value

            '--------------------------------------------------------------------------------------------------- 
            A_Main_Form.O_BOM_File_Name.Text = Exc.Worksheets("Sheet1").Range("A5").Value
            A_Main_Form.project_data_path.Text = Exc.Worksheets("Sheet1").Range("A9").Value
            A_Main_Form.re_o_bom_path.Text = Exc.Worksheets("Sheet1").Range("C3").Value

            '--------------------------------------------------------------------------------------------------- 프로그램 초기 실행시 "D:\STW\PROJECT_DB" 폴더가 없으면 자동 생성.
            Dim StrProject_DB_Path = Exc.Worksheets("Sheet1").Range("A" & 4).Value
            fso = CreateObject("Scripting.FileSystemObject")
            If StrProject_DB_Path <> "" Then
                If (fso.FolderExists(Exc.Worksheets("Sheet1").Range("A" & 4).Value)) = False Then
                    MkDir(Exc.Worksheets("Sheet1").Range("A" & 4).Value & "\")
                    Threading.Thread.Sleep(500)
                End If
            End If

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(False)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL")
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Find_data_Path 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


    '---------------------------------------------------------------------------------------------------
    ' Data_O_BOM_Excel_Search  &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '---------------------------------------------------------------------------------------------------
    Public Sub Data_O_BOM_Excel_Search()
        If date_value = "1" Then
            date_value = Format(Now, "yymmddhhmm")
        End If

        If now_o_bom_excel_path = "1" Then
            '--------------------------------------------------------------------------------------------------- now_o_bom_excel_path 가져오기
            Exc = CreateObject("excel.application")
            Exc.Workbooks.Open(Application.StartupPath & "\data_path.xlsx")

            now_o_bom_excel_path = Exc.Worksheets("Sheet1").Range("A" & 10).Value

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(False)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL")
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    Public Sub KillProcess(PName As String)
        ' Call KillExcel(Exc)
        Dim xlp() As Process = Process.GetProcessesByName(PName)
        Threading.Thread.Sleep(300)
        Dim i As Integer
        '---------------------------------------------------------------------------------------------------
        ' 생 성 일  : 20.03.20
        ' 생 성 자  : 허혜원
        ' 수 정 일  : 
        ' 수 정 자  : 
        ' 수정사유  : -
        ' Parameter : Kill Process 엑셀 종료 수정
        '            
        '---------------------------------------------------------------------------------------------------        
        If xlp.Length = 0 Then
            Exit Sub
        End If

        For i = xlp.Length - 1 To i <= 0 Step -1
            If i < 0 Then
                Exit For
            End If
            Dim Process As Process = xlp(i)
            Try
                'Process.CloseMainWindow()   
                'Process.Close()
                'Process.Refresh()
                If Process.StartTime >= ProgramStartDate Then
                    Process.Kill()  'Close와 차이는?
                End If
            Catch ex As System.ComponentModel.Win32Exception
                Debug.Print("The process is terimation or could not be terminated.")
            Catch ex As InvalidOperationException
                Debug.Print("The process has already exited.")
            Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                Debug.Print(ex.ToString() & " Exception caught.")
            End Try
        Next

        xlp = Process.GetProcessesByName(PName)
        Threading.Thread.Sleep(300)
        For Each Process As Process In xlp
            Try
                If Process.StartTime >= ProgramStartDate Then
                    Process.Close()
                    Process.Kill()
                End If
            Catch ex As System.ComponentModel.Win32Exception
                Debug.Print("The process is terimation or could not be terminated.")
            Catch ex As InvalidOperationException
                Debug.Print("The process has already exited.")
            Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                Debug.Print(ex.ToString() & " Exception caught.")
            End Try
        Next

    End Sub

    Public Sub KillExcel(killObject)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(killObject)
            killObject = Nothing
        Catch ex As Exception
            Debug.Print("A")
        End Try

    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Program_Version_Check : 실행된 레이아웃 프로그램 버전 검사
    '---------------------------------------------------------------------------------------------------
    Public Sub Program_Version_Check()
        Try
            Dim fso As FileSystemObject = New FileSystemObject
            '--------------------------------------------------------------------------------------------------- 현재 Data Version
            Dim oNowProgramVersion As File = fso.GetFile(Application.StartupPath & "\SM21_Layout_Program.exe")
            Dim oNowProgramModified As Date = oNowProgramVersion.DateLastModified

            '--------------------------------------------------------------------------------------------------- Server Data Version
            Dim oServerProgramVersion As File = fso.GetFile("Z:\WORK_REF\Layout_Program\SM21_Exe_File\SM21_Layout_Program.exe")
            Dim oServerProgramModified As Date = oServerProgramVersion.DateLastModified

            '--------------------------------------------------------------------------------------------------- Version이 다른 경우
            If Not oNowProgramModified = oServerProgramModified Then
                Call SHOW_MESSAGE_BOX("최신 Version 으로 Upgrade 하세요.", "", 1, 1)
                Version_Check_Value = "Old"
            Else
                Version_Check_Value = "New"
                Comp_Layout_Update_Value = "Last"
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("프로그램 업데이트 오류! 관리자에게 문의하세요.", "", 1, 1)
            A_Main_Form.Close()
            End
        End Try
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' O_BOM_Excel_Path : 폴더 경로를 엑셀에 저장
    '---------------------------------------------------------------------------------------------------
    Public Sub O_BOM_Excel_Path()
        '--------------------------------------------------------------------------------------------------- Excel Open 후 폴더 경로 저장
        Exc = CreateObject("excel.application")
        'Exc.Visible = True
        Exc.Workbooks.Open(Application.StartupPath & "\data_path.xlsx", 3, False)
        Exc.Worksheets("Sheet1").Range("A" & 5).Value = ""
        Exc.Worksheets("Sheet1").Range("A" & 5).Value = sO_BOM_File_Name
        Exc.Application.DisplayAlerts = False
        Exc.ActiveWorkbook.Close(False)
        Exc.Application.DisplayAlerts = True
        Exc.Quit()

        Call KillProcess("EXCEL")
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' Excle_Process_Count : 실행 된 PName의 프로세스 수
    '---------------------------------------------------------------------------------------------------
    Public Sub Excle_Process_Count(PName As String)
        Dim pgm = PName
        Dim wmi = GetObject("winmgmts:")
        Dim processes = wmi.execquery("select * from win32_process where name='" & pgm & "'")
        processes_count_value = processes.Count
        wmi = Nothing
    End Sub


    '---------------------------------------------------------------------------------------------------
    ' KillProcess_Layout_Program : 실행 된 레이아웃 프로그램이 있으면 삭제
    '---------------------------------------------------------------------------------------------------
    Public Sub KillProcess_Layout_Program()
        Try
            '--------------------------------------------------------------------------------------------------- Excel Close
            Threading.Thread.Sleep(200)

            Dim pgm As String = "SM21_Layout_Program.exe"
            Dim processes = GetObject("winmgmts:").execquery("select * from win32_process where name='" & pgm & "'")
            If processes.Count = 2 Then
                For Each process In processes
                    process.Terminate()
                    Exit For
                Next
            End If
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
        End Try
    End Sub


#Region " --- Base_Drawing_Create() : 기본 도면 설정 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.06.04
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Base_Drawing_Create()
        Try
            A_Main_Form.ProgressBar1.Value = 0

            '--------------------------------------------------------------------------------------------------- 최상위 Product를 ActiveDocument로 화면 변경
            For i = 1 To 10
                Call CATIA_BASIC02()

                If Right(CATIA.ActiveDocument.Name, 10) = "CATProduct" Then
                    Exit For
                End If

                CATIA.ActiveWindow.ActivateNext()
            Next

            '--------------------------------------------------------------------------------------------------- bullgear 찾기
            If bullgear_value Is Nothing Then
                Call CATIA_BASIC02()

                For i = 1 To ActiveDocument.Product.products.Count
                    If Right(ActiveDocument.Product.products.Item(i).Name, 9) = "EPN010101" Then
                        bullgear_value = ActiveDocument.Product.products.Item(i)
                        Exit For
                    End If
                Next
            End If

            If bullgear_value Is Nothing Then
                Call SHOW_MESSAGE_BOX("BullGear 정보가 없어 종료합니다.", "", 1, 1)
                Exit Sub
            End If

            '--------------------------------------------------------------------------------------------------- 도면Open
            drawingDocument1 = Documents.NewFrom(Application.StartupPath & "\COMP_Drawing.CATDrawing")
            drawing_selection = drawingDocument1.selection
            drawingSheets1 = drawingDocument1.sheets
            drawingSheet1 = drawingSheets1.Item("Model")
            drawingSheet1.ProjectionMethod = 1      ' catThirdAngle
            drawingViews1 = drawingSheet1.views
            A_Main_Form.ProgressBar1.Value = 20

            '--------------------------------------------------------------------------------------------------- Front View Axis 정의
            View_Axis_Value(1) = 1
            View_Axis_Value(2) = 0
            View_Axis_Value(3) = 0
            View_Axis_Value(4) = 0
            View_Axis_Value(5) = 0
            View_Axis_Value(6) = 1

            '--------------------------------------------------------------------------------------------------- Front View 생성
            Call Base_Drawing_View_Create(View_Axis_Value, 1, 130, 110, 0.03)
            A_Main_Form.ProgressBar1.Value = 40

            '--------------------------------------------------------------------------------------------------- Right View 생성
            Call Base_Drawing_View_Create(View_Axis_Value, 2, 300, 110, 0.03)
            A_Main_Form.ProgressBar1.Value = 60

            '--------------------------------------------------------------------------------------------------- Top View 생성
            Call Base_Drawing_View_Create(View_Axis_Value, 4, 130, 220, 0.03)
            A_Main_Form.ProgressBar1.Value = 80

            '--------------------------------------------------------------------------------------------------- ISO View 생성
            View_Axis_Value(1) = -0.707107
            View_Axis_Value(2) = 0.707107
            View_Axis_Value(3) = 0
            View_Axis_Value(4) = -0.408248
            View_Axis_Value(5) = -0.408248
            View_Axis_Value(6) = 0.816497

            Call Base_Drawing_View_Create(View_Axis_Value, 7, 300, 220, 0.015)
            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            ''--------------------------------------------------------------------------------------------------- Drawing 파일 저장
            CATIA.ActiveDocument.Update()
            CATIA.ActiveDocument.SaveAs(A_Main_Form.project_data_path.Text & "\LAYOUT\" & Sales_No & "_" & assy_value_Name & "_" & date_value & ".CATDrawing")
            A_Main_Form.ProgressBar1.Value = 100
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Base_Drawing_Create() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Base_Drawing_View_Create() : 기본 도면의 뷰 생성 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.06.04
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Base_Drawing_View_Create(View_Axis, Drawing_Option, X_Position, Y_Position, Drawing_Scale)
        Dim Base_Drawing_Create_View = Nothing
        Dim drawingViewGenerativeBehavior1 As DrawingViewGenerativeBehavior = Nothing
        Dim drawingViewGenerativeBehavior2 As DrawingViewGenerativeBehavior = Nothing
        Dim drawingViewGenerativeLinks1 As DrawingViewGenerativeLinks = Nothing
        Dim drawingViewGenerativeLinks2 As DrawingViewGenerativeLinks = Nothing



        If Drawing_Option = 1 Then
            '--------------------------------------------------------------------------------------------------- Front View 생성
            Base_Drawing_Front_View = drawingViews1.Add("AutomaticNaming")
            drawingViewGenerativeLinks1 = Base_Drawing_Front_View.GenerativeLinks
            drawingViewGenerativeBehavior1 = Base_Drawing_Front_View.GenerativeBehavior
            Dim productDocument1 = CATIA.Documents.Item(1)
            'Base_Drawing_Front_View.GenerativeBehavior.Document = product_add '191024
            drawingViewGenerativeBehavior1.Document = product_add



            '--------------------------------------------------------------------------------------------------- Product 정보가져오기
            Dim partDocument1 = Nothing
            For O_BOM_Num = 2 To O_BOM_Part_Value_Count - 1
                If O_BOM_Part_Value(O_BOM_Num, 3) = "EPN010101" Or O_BOM_Part_Value(O_BOM_Num, 3) = "EPN010101" Then
                    Dim O_BOM_Part_Name_Length = Len(O_BOM_Part_Value(O_BOM_Num, 1))

                    For kk = 1 To Documents.Count
                        If Left(Documents.Item(kk).Name, O_BOM_Part_Name_Length) = O_BOM_Part_Value(O_BOM_Num, 1) Then
                            partDocument1 = Documents.Item(kk)
                            Exit For
                        End If
                    Next kk

                    Exit For
                End If
            Next O_BOM_Num

            If partDocument1 Is Nothing Then
                Call SHOW_MESSAGE_BOX("O-BOM에 있는 BullGear Part 정보가 없습니다.", "", 1, 1)
                Exit Sub
            End If

            Dim axisSystem1 As AxisSystem = partDocument1.Part.AxisSystems.Item(1)
            axisSystem1.iscurrent = 1

            drawingViewGenerativeBehavior1.SetAxisSysteme(Nothing, axisSystem1)
            drawingViewGenerativeBehavior1.DefineFrontView(View_Axis(1), View_Axis(2), View_Axis(4), View_Axis(3), View_Axis(5), View_Axis(6)) '10 00 01

            Base_Drawing_Create_View = Base_Drawing_Front_View
        ElseIf Drawing_Option = 2 Or Drawing_Option = 3 Or Drawing_Option = 4 Or Drawing_Option = 5 Or Drawing_Option = 6 Then
            '-------------------------------------------------------------------------------------------------------- 새로운 view 생성
            drawingViewGenerativeBehavior1 = Base_Drawing_Front_View.GenerativeBehavior
            Base_Drawing_Create_View = drawingViews1.Add("AutomaticNaming")
            drawingViewGenerativeBehavior2 = Base_Drawing_Create_View.GenerativeBehavior

            '-------------------------------------------------------------------------------------------------------- 새로운 Right View 생성
            If Drawing_Option = 2 Then
                drawingViewGenerativeBehavior2.DefineProjectionView(drawingViewGenerativeBehavior1, 0)
                '-------------------------------------------------------------------------------------------------------- 새로운 Left View 생성
            ElseIf Drawing_Option = 3 Then
                drawingViewGenerativeBehavior2.DefineProjectionView(drawingViewGenerativeBehavior1, 1)
                '-------------------------------------------------------------------------------------------------------- 새로운 Top View 생성
            ElseIf Drawing_Option = 4 Then
                drawingViewGenerativeBehavior2.DefineProjectionView(drawingViewGenerativeBehavior1, 2)
                '-------------------------------------------------------------------------------------------------------- 새로운 Bottom View 생성
            ElseIf Drawing_Option = 5 Then
                drawingViewGenerativeBehavior2.RepresentationMode = Drawing_RepresentationMode_Value
                drawingViewGenerativeBehavior2.DefineProjectionView(drawingViewGenerativeBehavior1, 3)
                '-------------------------------------------------------------------------------------------------------- 새로운 Rear View 생성
            ElseIf Drawing_Option = 6 Then
                drawingViewGenerativeBehavior2.RepresentationMode = Drawing_RepresentationMode_Value
                drawingViewGenerativeBehavior2.DefineProjectionView(drawingViewGenerativeBehavior1, 4)
            End If

            drawingViewGenerativeLinks1 = Base_Drawing_Create_View.GenerativeLinks
            drawingViewGenerativeLinks2 = Base_Drawing_Front_View.GenerativeLinks
            drawingViewGenerativeLinks2.CopyLinksTo(drawingViewGenerativeLinks1)
        ElseIf Drawing_Option = 7 Then
            Base_Drawing_Create_View = drawingViews1.Add("AutomaticNaming")
            drawingViewGenerativeLinks1 = Base_Drawing_Create_View.GenerativeLinks
            drawingViewGenerativeBehavior1 = Base_Drawing_Create_View.GenerativeBehavior

            drawingViewGenerativeBehavior1.Document = product_add
            drawingViewGenerativeBehavior1.DefineFrontView(View_Axis(1), View_Axis(2), View_Axis(3), View_Axis(4), View_Axis(5), View_Axis(6))

            Base_Drawing_Create_View.Name = "ISO_View"
        End If

        Base_Drawing_Create_View.x = X_Position
        Base_Drawing_Create_View.y = Y_Position
        Base_Drawing_Create_View.scale2 = Drawing_Scale

        On Error Resume Next
        drawingViewGenerativeBehavior1.Update()
        Threading.Thread.Sleep(200)
        drawingViewGenerativeBehavior2.Update()
        Threading.Thread.Sleep(200)
        On Error GoTo 0

        drawing_selection.Clear()
        drawing_selection.Add(Base_Drawing_Create_View.Texts.Item(1))

        Dim drawingView_text
        If Not drawing_selection.Count = 0 Then
            drawingView_text = drawing_selection.Item(1).Value
            drawingView_text.Text = ""
        End If

        Base_Drawing_Create_View.Activate()
        drawingSheet1.Update()

        '--------------------------------------------------------------------------------------------------- Fit All In
        CATIA.ActiveWindow.ActiveViewer.Reframe()
    End Sub
#End Region


#Region " --- Ref_Drawing_Create() : 참고 도면 설정 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.06.04
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Ref_Drawing_Create()
        Try

            A_Main_Form.ProgressBar1.Value = 0

            '--------------------------------------------------------------------------------------------------- 최상위 Product를 ActiveDocument로 화면 변경
            For i = 1 To 10
                Call CATIA_BASIC02()

                If Right(CATIA.ActiveDocument.Name, 10) = "CATProduct" Then
                    Exit For
                End If

                CATIA.ActiveWindow.ActivateNext()
            Next

            '--------------------------------------------------------------------------------------------------- bullgear 찾기
            If bullgear_value Is Nothing Then
                Call CATIA_BASIC02()

                For i = 1 To ActiveDocument.Product.products.Count
                    If Right(ActiveDocument.Product.products.Item(i).Name, 9) = "EPN010101" Then
                        bullgear_value = ActiveDocument.Product.products.Item(i)
                        Exit For
                    End If
                Next
            End If

            If bullgear_value Is Nothing Then
                Call SHOW_MESSAGE_BOX("BullGear 정보가 없어 종료합니다.", "", 1, 1)
                Exit Sub
            End If

            '--------------------------------------------------------------------------------------------------- 화면 변경
            Call Active_Window_Change("CATProduct", 1, 10)

            Dim Data_Part_Name_Len
            '--------------------------------------------------------------------------------------------------- Enclosure 유무 확인
            Enclosure_Value = "Without Enclosure"
            Call Get_Result_Template_Item_Value("ENCLOSURE")
            selection.Clear()
            If Not Result_Template_Item_Value = "Not_Value" Then
                Data_Part_Name_Len = Len(Result_Template_Item_Value)
                For i = 1 To CATIA.ActiveDocument.Product.products.Count
                    If Right(CATIA.ActiveDocument.Product.products.Item(i).Name, Data_Part_Name_Len) = Result_Template_Item_Value Then
                        Enclosure_Value = "Hidden Enclosure"
                        '--------------------------------------------------------------------------------------------------- Enclosure Hide
                        selection.Add(CATIA.ActiveDocument.Product.products.Item(i))
                        selection.VisProperties.Parent.SetShow(1)
                        selection.Clear()
                        Exit For
                    End If
                Next
            End If

            '--------------------------------------------------------------------------------------------------- Coupling Cover Hide
            Call Get_Result_Template_Item_Value("COUPLING COVER")
            selection.Clear()
            If Not Result_Template_Item_Value = "Not_Value" Then
                Data_Part_Name_Len = Len(Result_Template_Item_Value)
                For i = 1 To CATIA.ActiveDocument.Product.products.Count
                    If Right(CATIA.ActiveDocument.Product.products.Item(i).Name, Data_Part_Name_Len) = Result_Template_Item_Value Then
                        '--------------------------------------------------------------------------------------------------- Enclosure Hide
                        selection.Add(CATIA.ActiveDocument.Product.products.Item(i))
                        selection.VisProperties.Parent.SetShow(1)
                        selection.Clear()
                        Exit For
                    End If
                Next
            End If
            A_Main_Form.ProgressBar1.Value = 20

            '--------------------------------------------------------------------------------------------------- 도면 add
            drawingDocument1 = CATIA.Documents.Add("Drawing")
            drawing_selection = drawingDocument1.selection
            drawingSheets1 = drawingDocument1.sheets
            drawingSheet1 = drawingSheets1.Item(1)
            drawingSheet1.ProjectionMethod = 1      ' catThirdAngle
            drawingViews1 = drawingSheet1.views

            '--------------------------------------------------------------------------------------------------- Front View Axis 정의
            View_Axis_Value(1) = -1
            View_Axis_Value(2) = 0
            View_Axis_Value(3) = 0
            View_Axis_Value(4) = 0
            View_Axis_Value(5) = 0
            View_Axis_Value(6) = 1

            '--------------------------------------------------------------------------------------------------- Front View 생성
            Call Base_Drawing_View_Create(View_Axis_Value, 1, 0, 0, 1)
            '--------------------------------------------------------------------------------------------------- Rigth View 생성
            Call Base_Drawing_View_Create(View_Axis_Value, 2, 8000, 0, 1)
            '--------------------------------------------------------------------------------------------------- Left View 생성
            Call Base_Drawing_View_Create(View_Axis_Value, 3, -8000, 0, 1)
            A_Main_Form.ProgressBar1.Value = 40

            '--------------------------------------------------------------------------------------------------- Top View 생성
            Call Base_Drawing_View_Create(View_Axis_Value, 4, 0, 5000, 1)
            '--------------------------------------------------------------------------------------------------- Bottom View 생성
            Call Base_Drawing_View_Create(View_Axis_Value, 5, 0, -5000, 1)
            '--------------------------------------------------------------------------------------------------- Rear View 생성
            Call Base_Drawing_View_Create(View_Axis_Value, 6, 8000, -5000, 1)
            A_Main_Form.ProgressBar1.Value = 60

            '--------------------------------------------------------------------------------------------------- ISO View 생성
            View_Axis_Value(1) = -0.707107
            View_Axis_Value(2) = 0.707107
            View_Axis_Value(3) = 0
            View_Axis_Value(4) = -0.408248
            View_Axis_Value(5) = -0.408248
            View_Axis_Value(6) = 0.816497
            Call Base_Drawing_View_Create(View_Axis_Value, 7, 8000, 5000, 1)
            A_Main_Form.ProgressBar1.Value = 80
            '--------------------------------------------------------------------------------------------------- Fit All In
            CATIA.ActiveWindow.ActiveViewer.Reframe()

            '--------------------------------------------------------------------------------------------------- Drawing 파일 저장
            CATIA.ActiveDocument.Update()

            If Enclosure_Value = "Without Enclosure" Then
                CATIA.ActiveDocument.SaveAs(A_Main_Form.project_data_path.Text & "\LAYOUT\" & Sales_No & "_" & assy_value_Name & "_" & _
                                            date_value & "_REF_W.CATDrawing")
            Else
                CATIA.ActiveDocument.SaveAs(A_Main_Form.project_data_path.Text & "\LAYOUT\" & Sales_No & "_" & assy_value_Name & "_" & _
                                            date_value & "_REF_H.CATDrawing")
            End If
            A_Main_Form.ProgressBar1.Value = 100
        Catch ex As Exception :  MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Ref_Drawing_Create() 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Get_Result_Template_Item_Value() : Template EPN No 가져오기 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.06.04
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Get_Result_Template_Item_Value(Item_Name)
        Result_Template_Item_Value = Nothing

        For i = 1 To 2

        Next

        For i = 1 To 10
            If STW_Template_Infomation(i, 1) = Item_Name Then
                Result_Template_Item_Value = STW_Template_Infomation(i, 2)
                Exit For
            End If
        Next
    End Sub
#End Region





End Module

