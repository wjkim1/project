Imports SPATypeLib
Imports NavigatorTypeLib
Imports System.Runtime.InteropServices
Imports System.Diagnostics '//190820
Imports System.Management
Imports System.Windows.Forms
Imports System.Windows

Module Common_Function
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2

    Private Declare Function SetWindowPos Lib "user32" (
        ByVal hwnd As Integer,
        ByVal hwndInsertAfter As Integer,
        ByVal x As Integer,
        ByVal y As Integer,
        ByVal cx As Integer,
        ByVal cy As Integer,
        ByVal wFlags As Integer
    ) As Long

    '----------------------------------------------------------------------- 간섭체크시 사용
    Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
    Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&

    '----------------------------------------------------------------------- 윈도우 창 찾기.
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function FindWindow(ByVal IpClassName As String,
                                 IpWindowName As String) As IntPtr
    End Function
    '----------------------------------------------------------------------- 찾은 윈도우 창 안에 세부내역 찾기.
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function FindWindowEx(ByVal parentHandle As IntPtr,
                                 ByVal childAfter As IntPtr,
                                 ByVal lclassName As String,
                                 ByVal windowTitle As String) As IntPtr
    End Function
    '----------------------------------------------------------------------- 이벤트 전달.
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function SendMessage(ByVal hWnd As IntPtr,
                                ByVal Msg As UInteger,
                                ByVal wParam As Integer,
                                ByVal lParam As Integer) As IntPtr
    End Function
    '----------------------------------------------------------------------- 
    Public BM_CLICK As Integer = &HF5


    '--------------------------------------------------------------------------------------------------- CATIA
    Public CATIA As INFITF.Application, Documents, product_add, products, ActiveDocument, selection, Exc, Is_F_S, CATIAProcessID
    '--------------------------------------------------------------------------------------------------- CATIA Drawing
    Public DrawingDocument, Drawingselection, DrawingSheets, DrawingSheet, DrawingViews, Main_View, Back_View, Back_Factory2D, drawing_error
    Public First_O_BOM, Diff_Path, Test_msg     'First_O_BOM = 1 O-BOM 가져오기

    Public Result_Part_Create(9) As Button
    Public Template_Name_Label(9) As Label
    Public Template_Create(9) As Button

    Public ProgramStartDate As Date
    '--------------------------------------------------------------------------------------------------- 현재 진행중인 Item Name
    Public Result_Template_Item_Name
    '--------------------------------------------------------------------------------------------------- 현재 진행중인 Item 값 Ex) 커플링 "EPN03"
    Public Result_Template_Item_Value
    '--------------------------------------------------------------------------------------------------- 현재 진행중인 Item 실적 Data 저장 Path
    Public Result_Template_Data_Path
    '--------------------------------------------------------------------------------------------------- 선택한 Form 순서 ( 선택한 Item Value 가져옴 )
    Public Form_Click_Index_Value
    '--------------------------------------------------------------------------------------------------- Help
    Public Help_Item_Name, Help_Item_Type
    '--------------------------------------------------------------------------------------------------- Version 정보
    Public Version_Check_Value, Comp_Layout_Update_Count
    '--------------------------------------------------------------------------------------------------- Error유무 확인
    Public Error_Count, Motor_BOM_Exist
    '--------------------------------------------------------------------------------------------------- BOM 파일명
    Public sO_BOM_File_Name
    '--------------------------------------------------------------------------------------------------- 간섭체크
    Public Clash_Check_EPN01, Clash_Check_EPN01_Name
    Public Clash_Check_EPN02, Clash_Check_EPN02_Name
    Public Clash_Check_Type, Clash_Check_Part_Count, Clash_Check_Distance_Value, Clash_Check_Excel_Count, Clash_Compute_Count
    Public Clash_Delete_Value(100), Check_Count

    '--------------------------------------------------------------------------------------------------- 프로그램에 필요한 기본 경로 변수화
    Public sDirectory_Path_Text, sConstraint_Excel_Path_Text, sO_BOM_Path_Text, sProject_DB_Path_Text, sTemplate_Data_Path_Text
    Public sLbl_ModelNumber

    '--------------------------------------------------------------------------------------------------- 선정시 축간 거리 값 변수
    Public axis_axis_measure_x, axis_axis_measure_y, axis_axis_measure_z
    '--------------------------------------------------------------------------------------------------- 템플릿 생성인지 선정인구 구분 변수
    Public iTemplateFlag As Integer = 0
    '--------------------------------------------------------------------------------------------------- 해당 파일이 있는지 검사용 변수
    Public sExcelFilePath As String



#Region " --- windows_ontop() : 현재 창을 윈도우 최상위 창으로 고정 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub windows_ontop(ByVal window As Long)
        SetWindowPos(window, HWND_TOPMOST, 0, 0, 0, 0, &H2 Or &H1)
    End Sub
#End Region


    Public Function RuntimeCheck()

        Dim counter As New System.Diagnostics.PerformanceCounter("Processor", "% Processor Time", "_Total")
        Dim Val As Decimal = counter.NextValue
        Threading.Thread.Sleep(500)
        Val = counter.NextValue
        Debug.Print("Runtime Check 1: " & Val.ToString())
        Threading.Thread.Sleep(1000)
        Val = counter.NextValue
        Debug.Print("Runtime Check 2: " & Val.ToString())


        Call Threading.Thread.Sleep(300)
        Dim hParent
        hParent = 0
        hParent = FindWindow(Nothing, "Runtime exception")

        If hParent <> 0 Then


            'Dim counter As PerformanceCounter


            MessageBox.Show("Runtime")
            Debug.Print("A")
        End If
    End Function

#Region " --- CATIA_BASIC() : CATIA의 실행 여부를 판단 및 문서 정보 갱신 ( 초기 Product가 없을때 ) "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Function CATIA_BASIC() As Object
        RuntimeCheck()
        CATIA = Nothing
        Try
            CATIA = GetObject(, "CATIA.Application")
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
        Finally
            If CATIA Is Nothing Then
                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
            End If
        End Try

        Documents = CATIA.Documents

        ''--------------------------------------------------------------------------------------------------- Message 제거
        'CATIA.DisplayFileAlerts = False

        'Dim ProcessList As System.Diagnostics.Process()
        'ProcessList = System.Diagnostics.Process.GetProcessesByName("CNEXT")

        'Dim Proc As System.Diagnostics.Process
        'For Each Proc In ProcessList
        '    If Proc.ProcessName = "CNEXT" Then
        '        CATIAProcessID = Proc.Id
        '    End If
        'Next

        Return 0
    End Function
#End Region


#Region " --- CATIA_BASIC02() : CATIA의 실행 여부를 판단 및 문서 정보 갱신 ( 초기 Product가 없을때 ) "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Function CATIA_BASIC02() As Object '6.0 Variant
        Try
            CATIA = GetObject(, "CATIA.Application")
            Documents = CATIA.Documents
            If Documents.Count <> 0 Then
                'If (Right(Documents.item(1).Name, 7) <> ".CATfct") Then
                If (Right(CATIA.ActiveDocument.Name, 7) <> ".CATfct") Then
                    ActiveDocument = CATIA.ActiveDocument
                    '--------------------------------------------------------------------------------------------------- 현재 ActiveDocument 가 Product일때만 Product 정보 가져오기
                    If TypeName(ActiveDocument) = "ProductDocument" Or TypeName(ActiveDocument) = "PartDocument" Then
                        Try
                            product_add = ActiveDocument.Product
                            products = product_add.products
                        Catch exx As Exception
                        End Try
                    End If
                    selection = CATIA.ActiveDocument.selection
                Else
                    MessageBox.Show("정상적으로 종료되지 않은 " & Documents.item(1).Name & "파일이 남아 있습니다. CATIA를 재실행해주세요. ")
                    Try
                        CATIA.DisplayFileAlerts = False
                        'Documents.item(1).Close
                        CATIA.DisplayFileAlerts = True
                    Catch exx As Exception
                    End Try
                End If
            End If
        Catch ex As Exception
            'MessageBox.Show("CATIA를 인식할 수 없습니다." & vbCrLf & "CATIA와 Layout 프로그램을 재실행해 주세요")
        End Try

        Return 0
    End Function
#End Region


#Region " --- CATIA_BASIC03() : CATIA의 실행 여부를 판단 및 문서 정보 갱신 ( 도면 ) "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Function CATIA_BASIC03() As Object
        Try
            DrawingDocument = CATIA.ActiveDocument
            Drawingselection = DrawingDocument.selection
            DrawingSheets = DrawingDocument.sheets
            DrawingSheet = DrawingSheets.ActiveSheet
            DrawingViews = DrawingSheet.views
            Main_View = DrawingViews.Item("Main View")
            Back_View = DrawingViews.Item("Background View")
            Back_Factory2D = Back_View.Factory2D
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
        End Try

        Return 0
    End Function
#End Region


#Region " --- O_BOM_Excel_Writing() : O-BOM 엑셀에 EPN Number, 선정 데이터 파일명 변경 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : EPN_No : 현재 아이템의 EPN
    '             File_Name : 파일명
    '---------------------------------------------------------------------------------------------------
    Public Sub O_BOM_Excel_Writing(EPN_No, File_Name)
        Try
            '--------------------------------------------------------------------------------------------------- Excel Open
            Exc = CreateObject("excel.application")
            'Exc.Visible = True
            Exc.Workbooks.Open(now_o_bom_excel_path)

            '--------------------------------------------------------------------------------------------------- Data 검토
            Dim now_o_bom_EPN_count As Integer = 1
            Dim Check As Boolean = True

            Do
                '--------------------------------------------------------------------------------------------------- EPN 이름 비교
                If EPN_No = CStr(Exc.Worksheets("BOM_LIST").Range("K" & now_o_bom_EPN_count + 1).Value) Then
                    Exc.Worksheets("BOM_LIST").Range("B" & now_o_bom_EPN_count + 1).Value = File_Name
                    Check = False
                    Exit Do
                End If

                If CStr(Exc.Worksheets("BOM_LIST").Range("B" & now_o_bom_EPN_count + 1).Value) = "" Then
                    Exc.Worksheets("BOM_LIST").Range("B" & now_o_bom_EPN_count + 1).Value = File_Name
                    Exc.Worksheets("BOM_LIST").Range("K" & now_o_bom_EPN_count + 1).Value = EPN_No
                    Check = False
                    Exit Do
                End If

                now_o_bom_EPN_count = now_o_bom_EPN_count + 1
            Loop Until Check = False

            Exc.Application.DisplayAlerts = False
            Exc.ActiveWorkbook.Close(True)
            Exc.Application.DisplayAlerts = True
            Exc.Quit()

            Call KillProcess("EXCEL")
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("O_BOM_Excel_Writing 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Template_Data_Form_Inicial_Setting() : 템플릿 생성 UI 초기화 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Sub Template_Data_Form_Inicial_Setting()
        Try
            '--------------------------------------------------------------------------------------------------- 현재 Tempate 구현 수량
            Dim Template_Form_Count = Result_Part_Create.Length - 1
            'Template_Form_Count = 9

            '--------------------------------------------------------------------------------------------------- 초기 Form Enable Setting
            For i = 1 To Template_Form_Count
                Result_Part_Create(i).Enabled = False
                Template_Name_Label(i).Text = ""
                'Template_Create(i) = False
                Template_Create(i).Enabled = False
            Next

            '--------------------------------------------------------------------------------------------------- Form Setting
            For i = 1 To STW_Template_Infomation_Count - 1
                '--------------------------------------------------------------------------------------------------- 실적형상 가져오기
                If Not (STW_Template_Infomation(i, 8) = "" Or STW_Template_Infomation(i, 8) = "N/A") Then
                    Result_Part_Create(i).Enabled = True
                End If
                '--------------------------------------------------------------------------------------------------- Template 생성
                If Not (STW_Template_Infomation(i, 9) = "" Or STW_Template_Infomation(i, 9) = "N/A") Then
                    Template_Create(i).Enabled = True
                End If
                '--------------------------------------------------------------------------------------------------- Template 이름 기입
                Template_Name_Label(i).Text = STW_Template_Infomation(i, 1)
            Next

            '   기    능 : O-BOM에 정보가 Base Frame or Simple Base 구분하여 UI에 버튼 활성화 변경
            '---------------------------------------------------------------------------------------------------
            If assy_value_Name = "SM3" Or assy_value_Name = "SM4" Or assy_value_Name = "SM5" Or assy_value_Name = "SM6" Then
                For i = 2 To O_BOM_Part_Value_Count - 1
                    If O_BOM_Part_Value(i, 3) = "EPN40" Then
                        For j = 1 To Template_Form_Count
                            If STW_Template_Infomation(j, 1) = "BASE FRAME" Then
                                Result_Part_Create(j).Enabled = False
                                Template_Create(j).Enabled = False
                            ElseIf STW_Template_Infomation(j, 1) = "SIMPLE BASE" Then
                                Result_Part_Create(j).Enabled = True
                                Template_Create(j).Enabled = False
                            End If
                        Next j
                    ElseIf O_BOM_Part_Value(i, 3) = "EPN18" Then
                        For j = 1 To Template_Form_Count
                            If STW_Template_Infomation(j, 1) = "BASE FRAME" Then
                                Result_Part_Create(j).Enabled = True
                                Template_Create(j).Enabled = True
                            ElseIf STW_Template_Infomation(j, 1) = "SIMPLE BASE" Then
                                Result_Part_Create(j).Enabled = False
                                Template_Create(j).Enabled = False
                            End If
                        Next j
                    End If
                Next i
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Template_Data_Form_Inicial_Setting 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Template_Data_Form_Inicial_Setting() : 선택한 버튼에 따라 실행 폼 설정 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 불필요한 소스 부분 삭제.
    ' Parameter : Form_Name : 선택한 버튼 이름
    '             Form_Option : 
    '---------------------------------------------------------------------------------------------------
    Public Sub Form_Loading_Function(Form_Name, Form_Option)
        Try
            '--------------------------------------------------------------------------------------------------- 선정시 다른 파트 지우는것 같아 초기화 / lsm / 190624
            Result_Template_Item_Value = ""

            '--------------------------------------------------------------------------------------------------- 현재 시간 및 O-BOM Path 가져오기
            Call Data_O_BOM_Excel_Search()

            Template_Data_Loading_Form.ProgressBar1_Change(20)

            If Not Form_Option = "Result_Data" Then
                Call Input_Output_Excel_Value(Form_Option)
            End If

            '--------------------------------------------------------------------------------------------------- 
            iTemplateFlag = 0       ' Axis 복붙시 생성 좌표값을 넣을지, 0으로 할지 정하는 변수 (1 이 좌표 측정해서 넣음)

            '--------------------------------------------------------------------------------------------------- Coupling 실적 형상
            If Form_Name = "B_Coupling_Create" Then
                B_Coupling_Create.Show()
            ElseIf Form_Name = "C_Coupling_Cover_Create" Then
                C_Coupling_Cover_Create.Show()
            ElseIf Form_Name = "D_Oil_Tank_Create" Then
                D_Oil_Tank_Create.Show()
            ElseIf Form_Name = "E_Oil_Piping_Create" Then
                E_Oil_Piping_Create.Show()
            ElseIf Form_Name = "F_AC_Outlet_Create" Then
                F_AC_Outlet_Create.Show()
            ElseIf Form_Name = "G_Enclosure_Create" Then
                G_Enclosure_Create.Show()
            ElseIf Form_Name = "H_Base_Frame_Create" Then
                H_Base_Frame_Create.Show()
                'ElseIf Form_Name = "I_Oil_Cooler_Adapter_Create" Then
                '    I_Oil_Cooler_Adapter_Create.Show
                'ElseIf Form_Name = "R_Simple_Base_Create" Then
                '    R_Simple_Base_Create.Show

                'ElseIf Form_Name = "L_Motor_Support_Create" Then
                '    L_Motor_Support_Create.Show
                'ElseIf Form_Name = "M_Installation_Jig_Create" Then
                '    M_Installation_Jig_Create.Show
                'ElseIf Form_Name = "N_Foundation_Create" Then
                '    N_Foundation_Create.Show
                'ElseIf Form_Name = "O_Oil_Return_Piping_Create" Then
                '    O_Oil_Return_Piping_Create.Show
                'ElseIf Form_Name = "Q_CABLE_Template_Create" Then
                '    Q_CABLE_Template_Create.Show

            ElseIf Form_Name = "CA_Coupling_Cover_Template_Create" Then
                iTemplateFlag = 1
                CA_Coupling_Cover_Template_Create.Show()
            ElseIf Form_Name = "DA_Oil_Tank_Template_Create" Then
                DA_Oil_Tank_Template_Create.Show()
            ElseIf Form_Name = "EA_Oil_Piping_Template_Create" Then
                EA_Oil_Piping_Template_Create.Show()
            ElseIf Form_Name = "FA_AC_Outlet_Template_Create" Then
                FA_AC_Outlet_Template_Create.Show()
            ElseIf Form_Name = "GA_Enclosure_Template_Create" Then
                GA_Enclosure_Template_Create.Show()
            ElseIf Form_Name = "HA_Base_Frame_Template_Create" Then
                HA_Base_Frame_Template_Create.Show()

                'ElseIf Form_Name = "LA_Motor_Support_Template_Create" Then
                '    LA_Motor_Support_Template_Create.Show
                'ElseIf Form_Name = "MA_Installation_Jig_Template_Create" Then
                '    MA_Installation_Jig_Template_Create.Show
                'ElseIf Form_Name = "NA_Foundation_Template_Create" Then
                '    NA_Foundation_Template_Create.Show
                'ElseIf Form_Name = "OA_Oil_Return_Piping_Template_Create" Then
                '    OA_Oil_Return_Piping_Template_Create.Show
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Form_Loading_Function 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Data_Constraint_Delete() : MSFlexGrid1_DblClick 이벤트시 기종_Constraint_List.xlsx 읽어 재구속 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : 불필요한 소스 부분 삭제.
    ' Parameter : Axis_Value : 
    '             Loop_Count : 
    '             Data_EPN_Number : 
    '---------------------------------------------------------------------------------------------------
    Public Sub Data_Constraint_Delete(Axis_Value, Loop_Count, Data_EPN_Number)
        Try
            Call CATIA_BASIC02()

            '--------------------------------------------------------------------------------------------------- Coupling Constraint Delete & Create
            For i = 1 To Loop_Count - 1
                Call Constraint_Delete(Axis_Value(i, 1), Axis_Value(i, 2), Axis_Value(i, 3))
            Next

            Dim Data_Part_Name_Len = Len(Data_EPN_Number)

            Call CATIA_BASIC02()
            selection.Clear()

            '--------------------------------------------------------------------------------------------------- EPN NUMBER 값이 없으면 엑셀 다시 읽기
            If Data_EPN_Number = "" Then
                Call Get_Excel_Value_Function(Application.StartupPath & "\STW_Template_Infomation.xlsx", assy_value_Name, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "",
                                              STW_Template_Infomation_Count, STW_Template_Infomation, "0")

                Result_Template_Item_Value = STW_Template_Infomation(Form_Click_Index_Value, 2)
                Data_EPN_Number = Result_Template_Item_Value
            End If

            '--------------------------------------------------------------------------------------------------- 구속조건 삭제 후 해당 파트 삭제
            If Not Data_EPN_Number = "" Then
                For i = 1 To CATIA.ActiveDocument.Product.products.Count
                    If Right(CATIA.ActiveDocument.Product.products.Item(i).Name, Data_Part_Name_Len) = Data_EPN_Number Then
                        selection.Add(CATIA.ActiveDocument.Product.products.Item(i))
                        selection.Delete()
                        selection.Clear()
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Data_Constraint_Delete 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Clash_Check_Execute() : 간섭체크시 실제 로직 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : Data_EPN_NO : 
    '---------------------------------------------------------------------------------------------------
    Public Sub Clash_Check_Execute(Data_EPN_NO)
        RuntimeCheck()

        Try
            '--------------------------------------------------------------------------------------------------- ZA_Clash_Check_Create Show
            ZA_Clash_Check_Create.EPN_NO_Label.Text = Data_EPN_NO
            ZA_Clash_Check_Create.Clash_Check_Create_Lebel.Text = "간섭체크 중..."
            '--------------------------------------------------------------------------------------------------- Clash_MSFlexGrid Clear
            ZA_Clash_Check_Create.DataGridView1.RowHeadersVisible = False
            ZA_Clash_Check_Create.DataGridView1.RowCount = 1
            '--------------------------------------------------------------------------------------------------- Clash_MSFlexGrid 초기 Setting
            ZA_Clash_Check_Create.DataGridView1.ColumnCount = 11
            ZA_Clash_Check_Create.DataGridView1.Columns(0).HeaderText = ""
            ZA_Clash_Check_Create.DataGridView1.Columns(1).HeaderText = "No"
            ZA_Clash_Check_Create.DataGridView1.Columns(2).HeaderText = "CHK EPN1 Name"
            ZA_Clash_Check_Create.DataGridView1.Columns(3).HeaderText = "CHK EPN2 Name"
            ZA_Clash_Check_Create.DataGridView1.Columns(4).HeaderText = "CHECK Result"
            ZA_Clash_Check_Create.DataGridView1.Columns(5).HeaderText = "간섭검증"
            ZA_Clash_Check_Create.DataGridView1.Columns(6).HeaderText = "CHK EPN 1"
            ZA_Clash_Check_Create.DataGridView1.Columns(7).HeaderText = "CHK EPN 2"
            ZA_Clash_Check_Create.DataGridView1.Columns(8).HeaderText = "CHECK Type"
            ZA_Clash_Check_Create.DataGridView1.Columns(9).HeaderText = "Min Distance"
            ZA_Clash_Check_Create.DataGridView1.Columns(10).HeaderText = "Req. Distance"
            ZA_Clash_Check_Create.DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(10).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            '--------------------------------------------------------------------------------------------------- ZA_Clash_Check_Create.Clash_MSFlexGrid Width
            ZA_Clash_Check_Create.DataGridView1.Columns(0).Width = 20
            ZA_Clash_Check_Create.DataGridView1.Columns(1).Width = 30
            ZA_Clash_Check_Create.DataGridView1.Columns(2).Width = 150
            ZA_Clash_Check_Create.DataGridView1.Columns(3).Width = 150
            ZA_Clash_Check_Create.DataGridView1.Columns(4).Width = 170
            ZA_Clash_Check_Create.DataGridView1.Columns(5).Width = 130
            ZA_Clash_Check_Create.DataGridView1.Columns(6).Width = 100
            ZA_Clash_Check_Create.DataGridView1.Columns(7).Width = 100
            ZA_Clash_Check_Create.DataGridView1.Columns(8).Width = 150
            ZA_Clash_Check_Create.DataGridView1.Columns(9).Width = 110
            ZA_Clash_Check_Create.DataGridView1.Columns(10).Width = 110

            '--------------------------------------------------------------------------------------------------- 1번 컬럼에 체크박스 추가
            Dim ChkBox As New DataGridViewCheckBoxColumn
            ChkBox.Name = "btnCheck"
            ChkBox.HeaderText = "Check"
            ZA_Clash_Check_Create.DataGridView1.Columns.Insert(0, ChkBox)
            ZA_Clash_Check_Create.DataGridView1.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(0).Width = 50

            Call CATIA_BASIC02()
            selection.Clear()
            RuntimeCheck()
            '--------------------------------------------------------------------------------------------------- Delete Grpup , Interference , Distance 지우기
            selection.search("Name=" & Data_EPN_NO & "C_*,all")
            If selection.Count > 0 Then
                Dim Clash_Delete_Count = 1
                For i = 1 To selection.Count
                    If selection.Item(i).Value.Parent.Name = "Group" Or selection.Item(i).Value.Parent.Name = "Interference" Or selection.Item(i).Value.Parent.Name = "Distance" Then
                        Clash_Delete_Value(Clash_Delete_Count) = selection.Item(i).Value
                        Clash_Delete_Count = Clash_Delete_Count + 1
                    End If
                Next

                selection.Clear()
                For i = 1 To Clash_Delete_Count - 1
                    selection.Add(Clash_Delete_Value(i))
                Next

                If Not selection.Count = 0 Then
                    selection.Delete()
                    selection.Clear()
                End If
            End If

            '-------------------------------------------------------------------------------------------------------------------Group 생성
            Dim cGroups = CATIA.ActiveDocument.Product.GetTechnologicalObject("Groups")
            '--------------------------------------------------------------------------------------------------- 계산
            Dim Clash_Check_Count = 1
            Clash_Compute_Count = 1
            '--------------------------------------------------------------------------------------------------- 가운데 정렬
            ZA_Clash_Check_Create.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ZA_Clash_Check_Create.DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            ZA_Clash_Check_Create.DataGridView1.RowCount = Clash_Check_Excel_Count - 1
            RuntimeCheck()
            Dim oGroup1 As Group, oGroup2 As Group
            Dim FirstGroup, SecondGroup
            For i = 1 To Clash_Check_Excel_Count - 1
                Dim Clash_Count = 0
                Dim Contact_Count = 0
                '-------------------------------------------------------------------------------------------------------------------새로운 Group01 생성
                oGroup1 = cGroups.AddFromSel
                oGroup1.Name = Data_EPN_NO & "C_" & Clash_Compute_Count & "_" & Clash_Check_EPN01(Clash_Check_Count)
                '-------------------------------------------------------------------------------------------------------------------새로운 Group02 추가
                oGroup2 = cGroups.Add
                oGroup2.Name = Data_EPN_NO & "C_" & Clash_Compute_Count & "_" & Clash_Check_EPN02(Clash_Check_Count)
                '-------------------------------------------------------------------------------------------------------------------새로운 Group01에 충돌 체크 Product 추가
                '---------------------------------------------------------------------------------------------------
                ' 생 성 일  : -
                ' 생 성 자  : -
                ' 수 정 일  : 19.08.22
                ' 수 정 자  : 허혜원
                ' 수정사유  : Product 내에 없는 Part 파일을 간섭체크 하는 경우 있어 오류 발생.
                '            간섭체크 전 ActiveDocument에 존재하는 Part인지 확인
                ' Parameter : -
                '---------------------------------------------------------------------------------------------------
                Dim IsItemExists
                IsItemExists = False
                For j = 1 To CATIA.ActiveDocument.Product.products.Count
                    If CATIA.ActiveDocument.Product.products.item(j).Name = Clash_Check_EPN01(Clash_Check_Count) Then
                        IsItemExists = True
                        Exit For
                    End If
                Next j
                If IsItemExists = True Then
                    oGroup1.AddExplicit(CATIA.ActiveDocument.Product.products.Item(Clash_Check_EPN01(Clash_Check_Count)))
                Else
                    Debug.Print(Clash_Check_EPN01(Clash_Check_Count) & " 항목 없어 간섭체크 안함")
                End If
                '--------------------------------------------------------------------------------------------------- Clash_MSFlexGrid
                ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(2).Value = Clash_Compute_Count
                ZA_Clash_Check_Create.DataGridView1.Rows(Clash_Check_Count - 1).Cells(7).Value = Clash_Check_EPN01(Clash_Check_Count)
                ZA_Clash_Check_Create.DataGridView1.Rows(Clash_Check_Count - 1).Cells(8).Value = Clash_Check_EPN02(Clash_Check_Count)
                ZA_Clash_Check_Create.DataGridView1.Rows(Clash_Check_Count - 1).Cells(9).Value = Clash_Check_Type(i)
                ZA_Clash_Check_Create.DataGridView1.Rows(Clash_Check_Count - 1).Cells(11).Value = Clash_Check_Distance_Value(Clash_Check_Excel_Count)

                '--------------------------------------------------------------------------------------------------- CHK EPN2 Group
                Dim k
                For j = 1 To Clash_Check_Part_Count(Clash_Check_Count)
                    If j > 1 Then
                        ZA_Clash_Check_Create.DataGridView1.RowCount = i + j
                    End If

                    '---------------------------------------------------------------------------------------------------
                    ' 생 성 일  : -
                    ' 생 성 자  : -
                    ' 수 정 일  : 19.08.22
                    ' 수 정 자  : 허혜원
                    ' 수정사유  : Product 내에 없는 Part 파일을 간섭체크 하는 경우 있어 오류 발생.
                    '            간섭체크 전 ActiveDocument에 존재하는 Part인지 확인
                    ' Parameter : -
                    '---------------------------------------------------------------------------------------------------
                    IsItemExists = False
                    For k = 1 To CATIA.ActiveDocument.Product.products.Count
                        If CATIA.ActiveDocument.Product.products.item(k).Name = Clash_Check_EPN02(Clash_Check_Count) Then
                            IsItemExists = True
                            Exit For
                        End If
                    Next k


                    Try
                        If IsItemExists = True Then
                            oGroup2.AddExplicit(CATIA.ActiveDocument.Product.products.Item(Clash_Check_EPN02(Clash_Check_Count)))
                        Else
                            Debug.Print(Clash_Check_EPN02(Clash_Check_Count) & " 항목 없어 간섭체크 안함")
                        End If
                        '--------------------------------------------------------------------------------------------------- Clash_MSFlexGrid
                        ZA_Clash_Check_Create.DataGridView1.Rows(Clash_Check_Count - 1).Cells(3).Value = Clash_Check_EPN01_Name(Clash_Check_Count)
                        ZA_Clash_Check_Create.DataGridView1.Rows(Clash_Check_Count - 1).Cells(4).Value = Clash_Check_EPN02_Name(Clash_Check_Count)
                        Clash_Check_Count = Clash_Check_Count + 1
                    Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
                        Clash_Check_Count += 1
                    End Try
                Next

                If Clash_Check_Type(i) = "CLASH + CONTACT" Or Clash_Check_Type(i) = "CLASH" Or Clash_Check_Type(i) = "CONTACT" Then
                    '------------------------------------------------------------------------------------------------------------------- Clash 추가
                    Dim Clashes As Clashes = CATIA.ActiveDocument.Product.GetTechnologicalObject("Clashes")
                    Dim oClash As Clash = Clashes.AddFromSel
                    '------------------------------------------------------------------------------------------------------------------- Clash Type 추가
                    oClash.ComputationType = 3
                    oClash.InterferenceType = 1
                    '------------------------------------------------------------------------------------------------------------------- Clash Selection01 추가
                    FirstGroup = oClash.FirstGroup
                    oClash.FirstGroup = oGroup1
                    '------------------------------------------------------------------------------------------------------------------- Clash Selection02 추가
                    SecondGroup = oClash.SecondGroup
                    oClash.SecondGroup = oGroup2
                    '------------------------------------------------------------------------------------------------------------------- Clash 계산
                    oClash.Compute()
                    oClash.Name = Data_EPN_NO & "C_" & Clash_Compute_Count & "_" & Clash_Check_EPN01(i) & "_" & Clash_Check_EPN02(i) & "_" & Clash_Check_Type(i)
                    '------------------------------------------------------------------------------------------------------------------- CLASH + CONTACT 일때...
                    If Clash_Check_Type(i) = "CLASH + CONTACT" Then
                        If oClash.Conflicts.Count = 0 Then
                            ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "Non CLASH , Non CONTACT"
                        Else
                            '------------------------------------------------------------------------------------------------------------------- Clash 계산
                            For j = 1 To oClash.Conflicts.Count
                                If oClash.Conflicts.Item(j).Type = 0 Then
                                    Clash_Count = Clash_Count + 1
                                ElseIf oClash.Conflicts.Item(j).Type = 1 Then
                                    Contact_Count = Contact_Count + 1
                                End If
                            Next
                            ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "CLASH : " & Clash_Count & " , CONTACT : " & Contact_Count
                        End If
                        '------------------------------------------------------------------------------------------------------------------- CLASH 일때...
                    ElseIf Clash_Check_Type(i) = "CLASH" Then
                        If oClash.Conflicts.Count = 0 Then
                            ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "Non CLASH"
                        Else
                            '------------------------------------------------------------------------------------------------------------------- Clash 계산
                            For j = 1 To oClash.Conflicts.Count
                                If oClash.Conflicts.Item(j).Type = 0 Then
                                    Clash_Count = Clash_Count + 1
                                End If
                            Next
                            If Clash_Count = 0 Then
                                ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "Non CLASH"
                            Else
                                ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "CLASH : " & Clash_Count
                            End If
                        End If
                        '------------------------------------------------------------------------------------------------------------------- CONTACT 일때...
                    ElseIf Clash_Check_Type(i) = "CONTACT" Then
                        If oClash.Conflicts.Count = 0 Then
                            ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "Non CONTACT"
                        Else
                            '------------------------------------------------------------------------------------------------------------------- Clash 계산
                            For j = 1 To oClash.Conflicts.Count
                                If oClash.Conflicts.Item(j).Type = 1 Then
                                    Contact_Count = Contact_Count + 1
                                End If
                            Next
                            If Contact_Count = 0 Then
                                ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "Non CONTACT"
                            Else
                                ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "CONTACT : " & Contact_Count
                            End If
                        End If
                    End If
                    '------------------------------------------------------------------------------------------------------------------- DISTANCE 일때...
                ElseIf Clash_Check_Type(i) = "DISTANCE" Then
                    ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(11).Value = Clash_Check_Distance_Value(i) & " mm"
                    '------------------------------------------------------------------------------------------------------------------- DISTANCE 추가
                    Dim cDistances As Distances = CATIA.ActiveDocument.Product.GetTechnologicalObject("Distances")
                    Dim oDistance As Distance = cDistances.Add
                    oDistance.ComputationType = 2
                    oDistance.FirstGroup = oGroup1
                    oDistance.SecondGroup = oGroup2
                    oDistance.Compute()
                    oDistance.Name = Data_EPN_NO & "C_" & Clash_Compute_Count & "_" & Clash_Check_EPN01(i) & "_" & Clash_Check_EPN02(i) & "_" & Clash_Check_Type(i)
                    '------------------------------------------------------------------------------------------------------------------- DISTANCE 결과
                    If Math.Round(oDistance.Value, 2) < Math.Round(Clash_Check_Distance_Value(i), 2) Then
                        ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "DISTANCE Fail"
                    Else
                        ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(5).Value = "DISTANCE Success"
                    End If
                    ZA_Clash_Check_Create.DataGridView1.Rows(i - 1).Cells(10).Value = Math.Round(oDistance.Value, 2) & " mm"
                End If

                i = Clash_Check_Count - 1
                Clash_Compute_Count = Clash_Compute_Count + 1
                RuntimeCheck()
            Next

            '------------------------------------------------------------------------------------------------------------------- 불필요한 컬럼 삭제
            ZA_Clash_Check_Create.DataGridView1.Columns.Remove(ZA_Clash_Check_Create.DataGridView1.Columns(1))

            ZA_Clash_Check_Create.Clash_Check_Create_Lebel.Text = "간섭체크 완료"
            ZA_Clash_Check_Create.Show()
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Clash_Check_Execute 함수 오류!!", "", 1, 1)
        End Try
    End Sub
#End Region


#Region " --- Get_Clash_Check_Excel_Value() : 간섭체크 결과를 엑셀파일에 저장 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : EPN_NO_Value : 
    '---------------------------------------------------------------------------------------------------
    Public Sub Get_Clash_Check_Excel_Value(EPN_NO_Value)
        '--------------------------------------------------------------------------------------------------- Excel Open 값 가져오기
        Exc = CreateObject("excel.application")
        'Exc.Visible = True

        sExcelFilePath = sConstraint_Excel_Path_Text & "\" & STW_COMP_Infomation(1, 4)
        If EXISTS_FILE_CHECK(sExcelFilePath) Then
            Exc.Workbooks.Open(sExcelFilePath)
        Else : Exit Sub
        End If

        '--------------------------------------------------------------------------------------------------- 배열 동적 할당
        ReDim Clash_Check_EPN01(0 To 100)
        ReDim Clash_Check_EPN02(0 To 100)
        ReDim Clash_Check_EPN01_Name(0 To 100)
        ReDim Clash_Check_EPN02_Name(0 To 100)
        ReDim Clash_Check_Type(0 To 100)
        ReDim Clash_Check_Part_Count(0 To 100)
        ReDim Clash_Check_Distance_Value(0 To 100)
        Clash_Check_Excel_Count = 1
        Dim Check = True
        Do   'Outer loop
            '--------------------------------------------------------------------------------------------------- Clash_Check_EPN01 가져오기
            Clash_Check_EPN01(Clash_Check_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("A" & Clash_Check_Excel_Count + 1).Value
            '--------------------------------------------------------------------------------------------------- Clash_Check_EPN02 가져오기
            Clash_Check_EPN02(Clash_Check_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("B" & Clash_Check_Excel_Count + 1).Value

            Clash_Check_EPN01_Name(Clash_Check_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("F" & Clash_Check_Excel_Count + 1).Value
            Clash_Check_EPN02_Name(Clash_Check_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("G" & Clash_Check_Excel_Count + 1).Value

            '--------------------------------------------------------------------------------------------------- Clash_Check_Type 가져오기
            Clash_Check_Type(Clash_Check_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("C" & Clash_Check_Excel_Count + 1).Value
            '--------------------------------------------------------------------------------------------------- Clash_Check_Part_Count 가져오기
            Clash_Check_Part_Count(Clash_Check_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("D" & Clash_Check_Excel_Count + 1).Value
            '--------------------------------------------------------------------------------------------------- Clash_Check_Distance_Value 가져오기
            Clash_Check_Distance_Value(Clash_Check_Excel_Count) = Exc.Worksheets(EPN_NO_Value).Range("E" & Clash_Check_Excel_Count + 1).Value
            Clash_Check_Excel_Count = Clash_Check_Excel_Count + 1
            '--------------------------------------------------------------------------------------------------- 조건 판별
            If Exc.Worksheets(EPN_NO_Value).Range("A" & Clash_Check_Excel_Count + 1).Value = "" Then        '조건이 True이면
                Check = False
                Exit Do
            End If
        Loop Until Check = False

        Exc.Application.DisplayAlerts = False
        Exc.ActiveWorkbook.Close(False)
        Exc.Application.DisplayAlerts = True
        Exc.Quit()

        Call KillProcess("EXCEL")

        '--------------------------------------------------------------------------------------------------- Count 입력 검증
        Dim Total_Count = 1
        Check_Count = 1
        For i = 1 To Clash_Check_Excel_Count - 1
            If Not CStr(Clash_Check_Part_Count(i)) = "" Then
                If Clash_Check_Part_Count(i) > 1 Then
                    For j = 1 To Clash_Check_Part_Count(i)
                        Total_Count = Total_Count + 1
                    Next
                ElseIf Clash_Check_Part_Count(i) = 1 Then
                    Total_Count = Total_Count + 1
                End If
            End If
            If i >= Total_Count Then
                Call SHOW_MESSAGE_BOX("간섭체크 Excel → " & EPN_NO_Value & " Sheet → CHECK_Count  값을 확인하세요.", "", 1, 1)
                Check_Count = 0
                Exit For
            End If
        Next
    End Sub
#End Region


#Region " --- Template_Clash_Check_Function() : Crash 실행 Function "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : - 
    '---------------------------------------------------------------------------------------------------
    Public Sub Template_Clash_Check_Function()
        '--------------------------------------------------------------------------------------------------- 간섭 체크 폴더 생성
        Call Clash_Check_Folder_Create()

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change(COMP_Product_File_Name, 0, 0)

        Try
            CATIA.ActiveDocument.Product.Update()
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
        End Try

        '--------------------------------------------------------------------------------------------------- Excel Value 가져오기
        Call Get_Clash_Check_Excel_Value(Result_Template_Item_Value)

        '--------------------------------------------------------------------------------------------------- Count Check
        If Check_Count = 0 Then Exit Sub

        '--------------------------------------------------------------------------------------------------- Clash 계산 및 결과값 입력
        Call Clash_Check_Execute(Result_Template_Item_Value)

        '--------------------------------------------------------------------------------------------------- 화면 변경
        Call Active_Window_Change(STW_Template_Infomation(Form_Click_Index_Value, 5), 0, 0)
    End Sub
#End Region


#Region " --- Active_Window_Change() : 카티아에 활성화 된 문서창 변경 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : Active_Document_Name : 변경 창
    '             Data_Name_Option : 변경 창 이름 +@
    '             Data_Name_Count : 열려있는 문서 창 개수
    '---------------------------------------------------------------------------------------------------
    Public Sub Active_Window_Change(Active_Document_Name, Data_Name_Option, Data_Name_Count)
        Dim WindowExistsCheck = False
        For i = 1 To CATIA.Windows.Count
            If InStr(CATIA.Windows.item(i).Caption, Active_Document_Name) <> 0 Then
                WindowExistsCheck = True
                Exit For
            End If
        Next i
        If WindowExistsCheck = False Then
            Exit Sub
        End If

        If Data_Name_Option = 0 Then
            Dim Name_Count = Len(Active_Document_Name)

            For i = 1 To 20
                Call CATIA_BASIC02()
                If Left(CATIA.ActiveDocument.Name, Name_Count) = Active_Document_Name Then
                    Exit For
                End If
                '--------------------------------------------------------------------------------------------------- 활성화 창을 다음 창으로 변환
                Threading.Thread.Sleep(2000)
                CATIA.ActiveWindow.ActivateNext()
                Threading.Thread.Sleep(2000)
            Next
        Else
            For i = 1 To 20
                Call CATIA_BASIC02()
                If Right(CATIA.ActiveDocument.Name, Data_Name_Count) = Active_Document_Name Then
                    Exit For
                End If
                '--------------------------------------------------------------------------------------------------- 활성화 창을 다음 창으로 변환
                Threading.Thread.Sleep(2000)
                CATIA.ActiveWindow.ActivateNext()
                Threading.Thread.Sleep(2000)
            Next
        End If
    End Sub
#End Region


#Region " --- RE_SELECTION_PART() : DataGridView1 를 더블 클릭하여 재 선정하기 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.08
    ' 수 정 자  : 이상민
    ' 수정사유  : 컨트롤 속성 값 변경에 따른 매개변수 값 변경
    ' Parameter : Form_Value : 폼 정보
    '             Constraint_Count : 생성되어 있는 구속조건의 총 개수
    '             Selection_Constraint : 생성 할 구속조건 정보
    '             Initial_Constraint : Axis 정보
    '             Initial_Constraint_Count : 구속 되어있는 개수
    '             Result_Template_Item_Value : 삭제 할 구속 조건의 EPN Number
    '---------------------------------------------------------------------------------------------------
    Public Sub RE_SELECTION_PART(Form_Value, Constraint_Count, Selection_Constraint, Initial_Constraint, Initial_Constraint_Count, Result_Template_Item_Value)
        Try
            Form_Value.ProgressBar1.Value = 0

            If Not CStr(Form_Value.DataGridView1.Rows(0).Cells(1).Value) = "" And Not CStr(Form_Value.DataGridView1.Rows(0).Cells(2).Value) = "" Then
                Form_Value.Message_Label.Text = "재 선정 작업 중입니다."
                Form_Value.ProgressBar1.Value = 10

                '--------------------------------------------------------------------------------------------------- 0번째행 ">" 삭제
                For i = 0 To Form_Value.DataGridView1.RowCount - 1
                    Form_Value.DataGridView1.Rows(i).Cells(0).Value = ""
                Next i

                Form_Value.ProgressBar1.Value = 20

                Dim msflexgrid_dbclick As Integer
                msflexgrid_dbclick = Form_Value.DataGridView1.SelectedRows(0).Index

                '--------------------------------------------------------------------------------------------------- 기 생성된 구속조건 및 Data 제거
                Call Data_Constraint_Delete(Initial_Constraint, Initial_Constraint_Count, Result_Template_Item_Value)
                'Call Data_Constraint_Delete(Coupling_Cover_Initial_Constraint, Coupling_Cover_Initial_Constraint_Count, Result_Template_Item_Value)
                Form_Value.ProgressBar1.Value = 30

                '--------------------------------------------------------------------------------------------------- Coupling_Cover Data open
                Dim arrayOfVariantOfBSTR1(0)
                arrayOfVariantOfBSTR1(0) = Form_Value.Data_Path_List.Items(msflexgrid_dbclick) & Form_Value.DataGridView1.Rows(msflexgrid_dbclick).Cells(2).Value

                If arrayOfVariantOfBSTR1(0).ToString.EndsWith(".CATPart") = False Then
                    arrayOfVariantOfBSTR1(0) = arrayOfVariantOfBSTR1(0) & ".CATPart"
                End If
                If Dir(arrayOfVariantOfBSTR1(0), vbDirectory) = "" Then

                    Dim fso
                    fso = CreateObject("Scripting.FileSystemObject")
                    Dim folder = fso.GetFolder(Form_Value.Data_Path_List.Items(msflexgrid_dbclick))
                    Dim GetFolderFiles = folder.Files
                    Dim FindFile = False
                    Dim FileName = ""
                    For Each EachFile In GetFolderFiles
                        If EachFile.Name.ToString().StartsWith(Form_Value.DataGridView1.Rows(msflexgrid_dbclick).Cells(2).Value) = True Then
                            FindFile = True
                            FileName = EachFile.Name
                            Exit For
                        End If
                    Next

                    '3D DB에 해당 파일이 없습니다.
                    If FindFile = False Then
                        MsgBox("3D DB에 해당 파일이 없습니다.", vbSystemModal)
                        Throw New System.Exception
                    Else
                        arrayOfVariantOfBSTR1(0) = Form_Value.Data_Path_List.Items(msflexgrid_dbclick) & FileName
                    End If
                End If
                products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
                products.Item(products.Count).Name = Result_Template_Item_Value
                'products.Item(Form_Value.DataGridView1.Rows(msflexgrid_dbclick).Cells(2).Value).Name = Result_Template_Item_Value
                '

                Form_Value.DataGridView1.Rows(msflexgrid_dbclick).Cells(0).Value = ">"
                Form_Value.ProgressBar1.Value = 50

                '--------------------------------------------------------------------------------------------------- 구속조건 삭제
                For i = 1 To Constraint_Count - 1
                    Call Constraint_Delete(Selection_Constraint(i, 1), Selection_Constraint(i, 2), 1)
                Next

                '--------------------------------------------------------------------------------------------------- 구속조건 생성
                For i = 1 To Constraint_Count - 1
                    Call Constraint_Delete(Selection_Constraint(i, 1), Selection_Constraint(i, 2), Selection_Constraint(i, 3))
                Next
                product_add.Update()

                '--------------------------------------------------------------------------------------------------- NOW O-BOM FILE DATA 기록
                Dim file_name_org = Split(Form_Value.DataGridView1.Rows(msflexgrid_dbclick).Cells(2).Value, ".CATPart")

                Form_Value.ProgressBar1.Value = 80

                Call O_BOM_Excel_Writing(Result_Template_Item_Value, file_name_org(0))
                Call Btn_Refresh_Click()

                Form_Value.ProgressBar1.Value = 100
                Form_Value.Message_Label.Text = "재 선정 작업 완료"
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Form_Value.Message_Label.Text = "재 선정 작업 중 오류 발생..."
            Form_Value.ProgressBar1.Value = 100
        End Try
    End Sub

#End Region


#Region " --- Generate_CATPart_from_Product() : Product -> Part 변환 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : -
    ' 생 성 자  : vb6에서 사용 된 함수
    ' 수 정 일  : 19.04.29
    ' 수 정 자  : 이상민
    ' 수정사유  : dll 사용 방법 변경 및 딜레이 추가
    ' Parameter : -
    '---------------------------------------------------------------------------------------------------
    Public Function Generate_CATPart_from_Product() As Object
        Dim iFlag As Integer = 0

        Call CATIA_BASIC02()

        Dim Temp_Product = Nothing
        Temp_Product = ActiveDocument.Product

        selection.Clear()
        selection.Add(Temp_Product)

        Threading.Thread.Sleep(300)
        CATIA.StartCommand("Generate CATPart from Product")
        Threading.Thread.Sleep(300)

        '--------------------------------------------------------------------------------------------------- 늦게 실행되는 경우를 위해..
        Dim hParent As IntPtr = 0
        Dim hChild As IntPtr = 0
        Dim bFlag As Boolean = False

        '--------------------------------------------------------------------------------------------------- 
        Do
            hParent = FindWindow(Nothing, "Generate CATPart from Product")

            If Not hParent = 0 Then
                bFlag = True
            Else
                iFlag += 1
                Threading.Thread.Sleep(200)

                If iFlag = 2000 Then
                    Return 0
                End If
            End If
        Loop While bFlag = False

        '--------------------------------------------------------------------------------------------------- 
        bFlag = False
        Do
            hChild = FindWindowEx(hParent, hChild, "Button", "OK")

            If Not hChild = 0 Then
                bFlag = True
                SendMessage(hChild, BM_CLICK, 0, 0)
            Else
                iFlag += 1
                Threading.Thread.Sleep(200)

                If iFlag = 2000 Then
                    Return 0
                End If
            End If
        Loop While bFlag = False

        Threading.Thread.Sleep(4000)

        Return 0
    End Function

#End Region


#Region " --- SHOW_MESSAGE_BOX() : 메세지 박스 최상위 출력을 위해.. "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : 19.05.03
    ' 생 성 자  : 이상민
    ' 수 정 일  : -
    ' 수 정 자  : -
    ' 수정사유  : -
    ' Parameter : sMessage : 출력 메세지
    '             iButton : 버튼 타입 정보
    '             iIcon : 아이콘 타입 정보
    '---------------------------------------------------------------------------------------------------
    Public Function SHOW_MESSAGE_BOX(ByVal sMessage As String, ByVal sTitle As String, ByVal iButton As Integer, ByVal iIcon As Integer) As System.Windows.Forms.DialogResult
        Try
            Dim sTitle2 As String
            If sTitle = "" Then
                sTitle2 = "ISPark Automation"
            Else
                sTitle2 = sTitle
            End If

            Dim oButtonType As System.Windows.Forms.MessageBoxButtons
            If iButton = 1 Then
                oButtonType = MessageBoxButtons.OK
            ElseIf iButton = 2 Then
                oButtonType = MessageBoxButtons.YesNo
            End If

            Dim oIconType As System.Windows.Forms.MessageBoxIcon
            If iIcon = 1 Then
                oIconType = MessageBoxIcon.Information
            ElseIf iIcon = 2 Then
                oIconType = MessageBoxIcon.Error
            End If

            If iButton = 2 Then
                Return MessageBox.Show(sMessage, sTitle2, oButtonType, oIconType, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            Else
                MessageBox.Show(sMessage, sTitle2, oButtonType, oIconType, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)   '터미네이터
                Return 0
            End If
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            MessageBox.Show(sMessage, "ISPark Automation", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return 0
        End Try
    End Function
#End Region


#Region " --- REMOVED_CLOSE_BUTTON() : 간섭체크를 위해 폼의 X버튼 비활성화 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : 19.05.08
    ' 생 성 자  : 이상민
    ' 수 정 일  : -
    ' 수 정 자  : -
    ' 수정사유  : -
    ' Parameter : iHWND : 폼 핸들 값
    '---------------------------------------------------------------------------------------------------
    Function REMOVED_CLOSE_BUTTON(ByVal iHWND As Integer) As Integer
        Dim iSysMenu As Integer
        iSysMenu = GetSystemMenu(iHWND, False)

        Return RemoveMenu(iSysMenu, SC_CLOSE, MF_BYCOMMAND)
    End Function
#End Region


#Region " --- EXISTS_FILE_CHECK() : 파일 존재 여부 판단 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : 19.05.10
    ' 생 성 자  : 이상민
    ' 수 정 일  : -
    ' 수 정 자  : -
    ' 수정사유  : -
    ' Parameter : sFullPath : 파일의 전체 경로
    '---------------------------------------------------------------------------------------------------
    Function EXISTS_FILE_CHECK(ByVal sFullPath As String) As Boolean
        If System.IO.File.Exists(sFullPath) Then
            Return True
        Else
            Call SHOW_MESSAGE_BOX("파일이 없어 작업을 취소합니다." & vbCrLf & "-> " & sFullPath, "", 1, 1)
            Return False
        End If
    End Function
#End Region


#Region " --- Result_Grid_Change() : 실적 검색 후 그리드에 표시된 데이터 변경시 데이터 경로의 폴더를 재 검색하여 파일명 가져오기 "
    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : 15.12.09
    ' 생 성 자  : 이상민
    ' 수 정 일  : 19.06.13
    ' 수 정 자  : 이상민
    ' 수정사유  : -
    ' Parameter : Form_Info : 적용 폼
    '             Grid_Selection_Number : 선택 된 그리드 행 넘버
    '---------------------------------------------------------------------------------------------------
    Public Function Result_Grid_Change(Form_Info, Grid_Selection_Number)
        Try
            '--------------------------------------------------------------------------------------------------- 폴더내 파일 검색
            Dim FName = Dir(sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\*.CATPart", vbNormal)

            'MsgBox(Form_Info.Controls.item("DataGridView1").Rows(0).Cells(2).Value)
            '--------------------------------------------------------------------------------------------------- 파일명 길이
            Dim FName_Len = Len(Form_Info.DataGridView1.Rows(Grid_Selection_Number).Cells(2).Value)
            'Dim FName_Len = Len(Form_Info.MSFlexGrid1.TextMatrix(Grid_Selection_Number, 2)) backup
            '--------------------------------------------------------------------------------------------------- 파일명 비교
            Do While FName <> ""
                If Left(FName, FName_Len) = Form_Info.DataGridView1.Rows(Grid_Selection_Number).Cells(2).Value Then
                    Call CATIA_BASIC02()

                    '--------------------------------------------------------------------------------------------------- 데이터 불러오기
                    Dim arrayOfVariantOfBSTR1(0)
                    arrayOfVariantOfBSTR1(0) = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\" & FName

                    products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
                    products.Item(products.Count).Name = Result_Template_Item_Value

                    'Form_Info.MSFlexGrid1.TextMatrix(Grid_Selection_Number, 0) = ">"
                    Exit Do
                End If

                '--------------------------------------------------------------------------------------------------- 다음 파일
                FName = Dir()
            Loop

            Return True
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Result_Grid_Change() 함수 오류!!", "", 1, 1)
            Return False
        End Try
    End Function
#End Region

    '---------------------------------------------------------------------------------------------------
    ' 생 성 일  : 20.02.04
    ' 생 성 자  : 허혜원
    ' 수 정 일  : 베이스 프레임 실적검색 후 오류 수정
    ' 수 정 자  : 
    ' 수정사유  : -
    ' Parameter : Form_Info : 적용 폼
    '             Grid_Selection_Number : 선택 된 그리드 행 넘버
    '---------------------------------------------------------------------------------------------------
    Public Function Result_Grid_Change_Base_Frame(Form_Info, Grid_Selection_Number)
        Try
            '--------------------------------------------------------------------------------------------------- 폴더내 파일 검색
            Dim FName = Dir(sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\*.CATPart", vbNormal)

            'MsgBox(Form_Info.Controls.item("DataGridView1").Rows(0).Cells(2).Value)
            '--------------------------------------------------------------------------------------------------- 파일명 길이
            Dim FName_Len = Len(Form_Info.Controls.item("DataGridView1").Rows(Grid_Selection_Number).Cells(2).Value)
            'Dim FName_Len = Len(Form_Info.MSFlexGrid1.TextMatrix(Grid_Selection_Number, 2)) backup
            '--------------------------------------------------------------------------------------------------- 파일명 비교
            Do While FName <> ""
                If Left(FName, FName_Len) = Form_Info.Controls.item("DataGridView1").Rows(Grid_Selection_Number).Cells(2).Value Then
                    Call CATIA_BASIC02()

                    '--------------------------------------------------------------------------------------------------- 데이터 불러오기
                    Dim arrayOfVariantOfBSTR1(0)
                    arrayOfVariantOfBSTR1(0) = sDirectory_Path_Text & "\" & Result_Template_Data_Path & "\" & FName

                    products.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
                    products.Item(products.Count).Name = Result_Template_Item_Value

                    'Form_Info.MSFlexGrid1.TextMatrix(Grid_Selection_Number, 0) = ">"
                    Exit Do
                End If

                '--------------------------------------------------------------------------------------------------- 다음 파일
                FName = Dir()
            Loop

            Return True
        Catch ex As Exception ': MessageBox.Show(ex.StackTrace.ToString) 'Debug.Print(ex.StackTrace)
            Call SHOW_MESSAGE_BOX("Result_Grid_Change() 함수 오류!!", "", 1, 1)
            Return False
        End Try
        End
    End Function

    Public Function SetControlBox(ByRef pform As Forms.Form, ByRef pCancel_button As Forms.Button, Show As Boolean)
        pform.ControlBox = Show
        pCancel_button.Enabled = Show
    End Function

End Module
