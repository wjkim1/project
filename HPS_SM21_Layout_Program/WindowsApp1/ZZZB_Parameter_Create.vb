Imports Scripting
Imports Microsoft.WindowsAPICodePack
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports INFITF
'---------------------------------------------------------------------------------------------------  Parameter 추가 및 삭제
' 생 성 일  : -
' 생 성 자  : -
' 수 정 일  : 1. 20.04.14
' 수 정 자  : 김원주
' 수정사유  : Visual Basic 6.0 -> Visual Studio2017 upgrade
' Parameter : -
'---------------------------------------------------------------------------------------------------


Public Class ZZZB_Parameter_Create
    Dim result_Value, sw, Job_Doc, Sql, Array_Size, Item_Name
    Dim file_number, Folder_Name, Origin_File_Name, File_Name(1000), Process_Count
    Dim Item_List, Selected_Item_Name, Undo_Item_Name, Item_Type
    Dim aaa, Split_Name, Test_Split_Name, i, One_folder_path, file_Attributes_value, strParam1
    Dim oShell, oFolder

    Private Sub SSTab2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SSTab2.SelectedIndexChanged


        If SSTab2.TabPages.Item(0).Visible = True Then
            Chk_Option.CheckState = False
            SSTab2.Height = 6735
            Me.Height = 659

        ElseIf SSTab2.TabPages.Item(1).Visible = True Then
            SSTab2.Height = 2655
            Me.Height = 380
        End If
    End Sub

    Private Sub Btn_Clear_Click(sender As Object, e As EventArgs) Handles Btn_Clear.Click
        List_Box.Items.Clear()
        Lbl_Path.Text = "-"
        Lbl_State.Text = "-"
        ProgressBar3.Value = 0
    End Sub

    Private Sub Btn_Folder_Click(sender As Object, e As EventArgs) Handles Btn_Folder.Click
        ProgressBar3.Value = 0
        Process_Count = 3
        '--------------------------------------------------------------------------------------------------- 폴더 선택
        List_Box.Items.Clear()
        Dim fso, floder, file, files, fileNameList()
        '---------------------------------------------------------------------------------------------------
        '   생 성 일 : 15-09-09
        '   생 성 자 : 이연중
        '   수 정 일 :
        '   수 정 자 :
        '   수정사유 :
        '   기    능 : 폴더 선택 에러 무시
        '   Parameter:
        '---------------------------------------------------------------------------------------------------
        'On Error Resume Next
        '---------------------------------------------------------------------------------------------------  Parameter 추가 및 삭제
        ' 생 성 일  : 20-04-14
        ' 생 성 자  : -
        ' 수 정 일  : -
        ' 수 정 자  : 김원주
        ' 수정사유  : FileBrowgerDialog 사용
        ' Parameter : -
        '---------------------------------------------------------------------------------------------------
        FolderDialog1.Description = "Folder  선택"
        FolderDialog1.ShowDialog()

        fso = CreateObject("Scripting.FileSystemObject")
        oFolder = FolderDialog1.SelectedPath


        StrFolder = Len(oFolder)
        Lbl_Path.Text = oFolder '& "\"
        Call lsGetFOS(oFolder & "\", 0)
        'On Error GoTo 0
    End Sub

    Private Sub lsGetFOS(lsPath As String, lsLeve As Long)
        Dim fso As Object
        Dim folder As Object
        Dim folders As Object
        Dim fc As Object
        Dim f, f1
        Dim lsCnt As Long
        Dim lsFolder As String

        lsCnt = lsLeve + 1
        fso = CreateObject("Scripting.FileSystemObject")
        folder = fso.getfolder(lsPath & "\")
        folders = folder.SubFolders

        fc = folder.files
        For Each f1 In fc
            'Text1 = Text1 & Space(lsLeve * 3) & folder.path & "\" & f1.Name & vbCrLf
            If Strings.Right(f1.Name, 8) = ".CATPart" Then

                List_Box.Items.Add(Microsoft.VisualBasic.Strings.Right(f1.Path, Len(f1.Path) - StrFolder))
            End If
        Next

        'Screen.MousePointer = 11
        For Each folder In folders

            If (GetAttr(folder) And vbDirectory) = vbDirectory Then
                f = fso.getfolder(folder.path)

                lsFolder = folder
                Call lsGetFOS(lsFolder, lsCnt)
            End If
        Next folder

        ' Screen.MousePointer = 1
        'Me.MousePointer = vbDefault
    End Sub


    Private Sub Btn_Members_Click(sender As Object, e As EventArgs) Handles Btn_Members.Click
        ProgressBar3.Value = 0
        Process_Count = 2
        '--------------------------------------------------------------------------------------------------- 파일 경로 가져오기
        List_Box.Items.Clear()
        '--------------------------------------------------------------------------------------------------- File Select
        CommonDialog1.InitialDirectory = "D:\STW\"
        CommonDialog1.FileName = "Open File"
        CommonDialog1.Filter = "CATIAFile|*.CATPart"
        CommonDialog1.ShowDialog()

        '--------------------------------------------------------------------------------------------------- 선택한 파일이 없으면.
        If CommonDialog1.FileName = "" Then
            aaa = MsgBox("파일을 선택하여 주세요. 종료합니다.", vbInformation, "ISPark_Automation")
            Exit Sub
        End If
        '--------------------------------------------------------------------------------------------------- List Box에 선택한 파일명 추가.
        sfile = Split(CommonDialog1.FileName, vbNullChar)

        If UBound(sfile) = 0 Then
            sfile = Split(CommonDialog1.FileName, "\")

            Lbl_Path.Text = sfile(0) & "\"

            For i = 1 To UBound(sfile) - 1
                Lbl_Path.Text = Lbl_Path.Text & sfile(i) & "\"
            Next

            List_Box.Items.Add(sfile(i))
        Else
            Lbl_Path.Text = sfile(0) & "\"

            For i = 1 To UBound(sfile)
                List_Box.Items.Add(sfile(i))
            Next i
        End If
    End Sub

    Dim Active_documents, partDocument, parameters1, Job_Document, reference1, publication1


    Private Sub Chk_Option_CheckedChanged_1(sender As Object, e As EventArgs) Handles Chk_Option.CheckedChanged
        Call Chk_Option_Click()
    End Sub

    Private Sub Cmd_OK_Click(sender As Object, e As EventArgs) Handles Cmd_OK.Click
        '---------------------------------------------------------------------------------------------------
        aaa = MsgBox("Renaming 기능을 실행 하겠습니까?", vbOKCancel Or vbInformation, "CATIA Automation")
        If aaa = 2 Then Exit Sub

        Read_Only = 0
        fso = CreateObject("Scripting.FileSystemObject")
        Lbl_State.Text = "이름 변경 중..."
        ProgressBar3.Value = 0
        '---------------------------------------------------------------------------------------------------
        If Txt_Found.Text = "" Then
            aaa = MsgBox("찾을 내용의 값을 입력해주세요.", vbInformation, "ISPark_Automation")
            Exit Sub
            '---------------------------------------------------------------------------------------------------
        Else
            If Not Process_Count = 1 And List_Box.Items.Count = 0 Then
                aaa = MsgBox("File List에 값을 입력해주세요.", vbInformation, "ISPark_Automation")
                Exit Sub
            End If
            '---------------------------------------------------------------------------------------------------
            Call Init_Chk_Array()
            ProgressBar3.Value = 15
            '--------------------------------------------------------------------------------------------------- 변경 버튼 클릭.
            On Error Resume Next
            If Process_Count = 1 Then
                Call Cmd_OK_Process()

                CATIA.ActiveDocument.Save()
                'CATIA.ActiveDocument.Close
                '--------------------------------------------------------------------------------------------------- Members 버튼 클릭 or Folder 버튼 클릭
            Else
                ProgressBar3.Value = 25

                For i = 0 To List_Box.Items.Count - 1
                    sfile = Lbl_Path.Text & List_Box.Items.Item(i)
                    '--------------------------------------------------------------------------------------------------- 읽기 전용 파일인지 검사.
                    file_Attributes_value = fso.GetFile(sfile)
                    'Call Delay(0.5)

                    If file_Attributes_value.Attributes = 33 Then
                        file_Attributes_value.Attributes = 32                   '읽기전용 X
                        'file_Attributes_value.Attributes = 33                   '읽기전용 O
                        Read_Only = 1
                    End If

                    On Error Resume Next

                    Job_Document = Documents.Open(sfile)

                    Call Cmd_OK_Process()

                    CATIA.ActiveDocument.Save()
                    CATIA.ActiveDocument.Close()
                    '--------------------------------------------------------------------------------------------------- 원본이 읽기 전용이였으면 읽기 전용으로 바꿔주기.
                    If Read_Only = 1 Then
                        'file_Attributes_value.Attributes = 32                   '읽기전용 X
                        file_Attributes_value.Attributes = 33                   '읽기전용 O
                    End If

                    ProgressBar3.Value = ProgressBar3.Value + (70 / List_Box.Items.Count)
                    On Error GoTo 0
                Next i
            End If
            On Error GoTo 0
        End If

        ProgressBar3.Value = 100
        Lbl_State.Text = "이름 변경 완료."
    End Sub

    Public Function Cmd_OK_Process()
        On Error Resume Next
        Dim Item_Name, temp, Selected_Item_Length, Found_Length, Replace_Length
        Dim Left_Name, Rigth_Name, Result_Item_Name
        '---------------------------------------------------------------------------------------------------
        Call CATIA_BASIC02()
        '--------------------------------------------------------------------------------------------------- Txt_Found.Text의 내용으로 검색.
        selection.Clear
        CATIA.ActiveDocument.Product.Update
        selection.search("Name=*" & Txt_Found.Text & "*,all")

        ReDim Item_List(selection.Count)                                    '선택한 대상 이름
        ReDim Selected_Item_Name(selection.Count)                           '작업 대상 이름
        ReDim Undo_Item_Name(selection.Count)                               '작업 전 원래 이름

        '---------------------------------------------------------------------------------------------------
        '   생 성 일 : 15.09.
        '   생 성 자 : 허혜원
        '   수 정 일 :
        '   수 정 자 :
        '   수정사유 :
        '   기    능 : 특정 타입일 경우 이름만 변경
        '   Parameter:
        '---------------------------------------------------------------------------------------------------
        For Job_Count = 1 To selection.Count
            If selection.Item(Job_Count).Type = "Add" Or selection.Item(Job_Count).Type = "Assemble" Or selection.Item(Job_Count).Type = "Body" Or selection.Item(Job_Count).Type = "HybridBody" Then
                Item_Name = selection.Item(Job_Count).Value.Name
                temp = Replace(Item_Name, Txt_Found.Text, Txt_Replace.Text)
                selection.Item(Job_Count).Value.Name = temp
            Else
                '--------------------------------------------------------------------------------------------------- 작업 대상 설정.
                Call Selection_Division(Item_List, Selected_Item_Name, Undo_Item_Name, Job_Count, Item_Type)

                If Not Item_Type = 0 Then
                    Selected_Item_Length = Len(Selected_Item_Name(Job_Count))       '선택한 대상의 이름 길이.
                    Found_Length = Len(Txt_Found.Text)                              '찾을 내용의 이름 길이.
                    Replace_Length = Len(Txt_Replace.Text)                          '변경할 내용의 이름 길이.

                    '--------------------------------------------------------------------------------------------------- 선택한 대상의 이름에서 찾을 내용의 이름 위치 검색
                    For i = 1 To Selected_Item_Length
                        If UCase(Mid(Selected_Item_Name(Job_Count), i, Found_Length)) = UCase(Txt_Found.Text) Then
                            Left_Name = Strings.Left(Selected_Item_Name(Job_Count), i - 1)
                            Rigth_Name = Mid(Selected_Item_Name(Job_Count), i + Found_Length, Selected_Item_Length)
                            Result_Item_Name = Left_Name & Txt_Replace.Text & Rigth_Name

                            '--------------------------------------------------------------------------------------------------- 선택한 대상의 종류에 따라 이름 변경.
                            If Item_Type = 1 Or Item_Type = 3 Then
                                Item_List(Job_Count).Name = Result_Item_Name
                                Exit For
                            ElseIf Item_Type = 2 Then
                                If Strings.Right(ActiveDocument.Name, 10) = "CATProduct" Then
                                    Job_Doc = Item_List(Job_Count).Parent.Parent.Parent.Part.Parameters
                                Else
                                    Job_Doc = ActiveDocument.Part.Parameters
                                End If

                                Call Job_Doc.Item(Selected_Item_Name(Job_Count)).Rename(Result_Item_Name)
                                Exit For
                            End If
                        End If
                    Next i
                End If
            End If
        Next Job_Count

        If Chk_Option_2.CheckState = 1 Then
            Lbl_State.Text = "Publication 재 생성 중..."
            Call Create_Publication()
            Lbl_State.Text = "Publication 생성 완료."
        End If

        On Error GoTo 0
    End Function
    '--------------------------------------------------------------------------------------------------- 선택한 요소의 종류 찾는 함수.
    Public Function Selection_Division(Item_List, Selected_Item_Name, Undo_Item_Name, Job_Count, ByRef Item_Type)
        On Error Resume Next

        Item_List(Job_Count) = selection.Item(Job_Count).Value
        Item_Type = selection.Item(Job_Count).Type

        If Item_Type = "Empty" Then
            Item_Type = 0
        End If
        '---------------------------------------------------------------------------------------------------
        Array_Size = UBound(Chk_Array)
        '---------------------------------------------------------------------------------------------------
        For i = 1 To Array_Size
            'If Chk_Array(i) = Item_Type Or Chk_Array(i) = Strings.Right(Item_Type, 5) Or Chk_Array(i) = Mid(Item_Type, 12, 5) Or Chk_Array(i) = Strings.Left(Item_Type, 5) Then
            If Chk_Array(i) = Item_Type Or Chk_Array(i) = Microsoft.VisualBasic.Strings.Right(Item_Type, 5) Or Chk_Array(i) = Mid(Item_Type, 12, 5) Or Chk_Array(i) = Strings.Left(Item_Type, 5) Then
                If Microsoft.VisualBasic.Strings.InStr(Chk_Array(i), Item_Type) <> 0 Then
                    Stop
                End If

                'If Chk_Array(i) = Item_Type Or Chk_Array(i) = Right(Item_Type, 5) Or Chk_Array(i) = Mid(Item_Type, 12, 5) Or Chk_Array(i) = Left(Item_Type, 5) Or Item_Type = "Add" Or Item_Type = "Assemble" Or Item_Type = "Body" Or Item_Type = "HybridBody" Then
                '            If Item_Type = "Add" Or Item_Type = "Assemble" Or Item_Type = "Body" Or Item_Type = "HybridBody" Then
                '                Item_Name = Item_List(Job_Count).Name
                '                temp = Replace(Item_Name, Txt_Found.Text, Txt_Replace.Text)
                '                Item_List(Job_Count).Name = temp
                '            End If

                If Item_Type = "AxisSystem" Then
                    Item_Name = Item_List(Job_Count).Name
                    Item_Type = 1
                    Exit For
                    'ElseIf Strings.Right(Item_Type, 5) = "Param" Or Item_Type = "Length" Then
                ElseIf Microsoft.VisualBasic.Strings.Right(Item_Type, 5) = "Param" Or Item_Type = "Length" Then
                    Item_Name = Split(Item_List(Job_Count).Name, "\")
                    Item_Name = Item_Name(UBound(Item_Name))
                    Item_Type = 2
                    Exit For
                ElseIf Strings.Left(Item_Type, 11) = "HybridShape" Then
                    Item_Name = Item_List(Job_Count).Name
                    Item_Type = 3
                    Exit For
                Else
                    Item_Name = Item_List(Job_Count).Name
                    Item_Type = 1
                End If
            End If
        Next i

        If Not Item_Type = 1 And Not Item_Type = 2 And Not Item_Type = 3 Then
            Item_Type = 0
            Exit Function
        End If
        '--------------------------------------------------------------------------------------------------- 실행 취소 기능을 위하여 따로 저장.
        Selected_Item_Name(Job_Count) = Item_Name
        Undo_Item_Name(Job_Count) = Item_Name

        On Error GoTo 0
    End Function

    '---------------------------------------------------------------------------------------------------
    '   Publication 만들기.
    '---------------------------------------------------------------------------------------------------
    Public Function Create_Publication()
        On Error Resume Next

        publications1 = CATIA.ActiveDocument.Product.publications
        '--------------------------------------------------------------------------------------------------- Publication 전부 삭제.
        Do
            For i = 1 To publications1.Count
                publications1.Remove(publications1.Item(i).name)

                'Call Delay(0.5)
            Next i
        Loop Until publications1.Count = 0

        Sql = ""
        '--------------------------------------------------------------------------------------------------- 변경 규칙의 Curve
        If Not TXT_Curve.Text = "" Then
            Sql = "CATPrtSearch.Curve.Name=*" + TXT_Curve.Text & "*"
        End If
        '--------------------------------------------------------------------------------------------------- 변경 규칙의 Axis1
        If Not TXT_Curve.Text = "" And Not TXT_Axis1.Text = "" Then
            Sql = Sql + " + CATPrtSearch.AxisSystem.Name=*" + TXT_Axis1.Text & "*"
        ElseIf Not TXT_Axis1.Text = "" Then
            Sql = "CATPrtSearch.AxisSystem.Name=*" + TXT_Axis1.Text & "*"
        End If
        '--------------------------------------------------------------------------------------------------- 변경 규칙의 Axis2
        If (Not TXT_Curve.Text = "" Or Not TXT_Axis1.Text = "") And Not TXT_Axis2.Text = "" Then
            Sql = Sql + " + CATPrtSearch.AxisSystem.Name=*" + TXT_Axis2.Text & "*"
        ElseIf Not TXT_Axis2.Text = "" Then
            Sql = "CATPrtSearch.AxisSystem.Name=*" + TXT_Axis2.Text & "*"
        End If
        Dim Selection_Val, TEMP_Name
        Dim Replace_Name
        '--------------------------------------------------------------------------------------------------- Publication 대상 찾기.
        'selection.search "(CATPrtSearch.Curve.Name=*" & TXT_Curve.Text & "* + CATPrtSearch.AxisSystem.Name=*" & TXT_Axis1.Text & "* + CATPrtSearch.AxisSystem.Name=*" & TXT_Axis2.Text & "*),all"
        selection.search("(" & Sql & "),all")
        '--------------------------------------------------------------------------------------------------- Publication 생성


        For i = 1 To selection.Count
            Selection_Val = selection.Item(i).Value

            TEMP_Name = Selection_Val.Name
            Selection_Val.Name = "TEMP_Axis"

            '
            add_axis_Product = ActiveDocument '.GetItem(ActiveDocument.Name)
            '   " & add_axis_Product.name/!TEMP_Axis"
            Replace_Name = Replace(add_axis_Product.name, ".CATPart", "")
            reference1 = add_axis_Product.Product.CreateReferenceFromName(" & Replace_Name/!TEMP_Axis")

            Selection_Val.Name = TEMP_Name

            publication1 = publications1.Add(TEMP_Name)

            publications1.SetDirect(TEMP_Name, reference1)

        Next i

        On Error GoTo 0
    End Function


    '---------------------------------------------------------------------------------------------------
    '   검색 조건 만들기.
    '---------------------------------------------------------------------------------------------------
    Public Function Init_Chk_Array()

        result_Value = 1
        sw = 1

        If Chk_2.CheckState = 1 Or Chk_1.CheckState = 1 Then result_Value = result_Value + 1
        If Chk_3.CheckState = 1 Or Chk_1.CheckState = 1 Then result_Value = result_Value + 1
        If Chk_4.CheckState = 1 Or Chk_1.CheckState = 1 Then result_Value = result_Value + 1
        If Chk_5.CheckState = 1 Or Chk_1.CheckState = 1 Then result_Value = result_Value + 1
        If Chk_6.CheckState = 1 Or Chk_1.CheckState = 1 Then result_Value = result_Value + 1
        If Chk_7.CheckState = 1 Or Chk_1.CheckState = 1 Then result_Value = result_Value + 1

        ReDim Chk_Array(result_Value)

        If Chk_2.CheckState = 1 Or Chk_1.CheckState = 1 Then
            Chk_Array(sw) = "Param"
            sw = sw + 1
        End If

        If Chk_3.CheckState = 1 Or Chk_1.CheckState = 1 Then
            Chk_Array(sw) = Chk_3.Text
            sw = sw + 1
        End If

        If Chk_4.CheckState = 1 Or Chk_1.CheckState = 1 Then
            Chk_Array(sw) = "Point"
            sw = sw + 1
        End If

        If Chk_5.CheckState = 1 Or Chk_1.CheckState = 1 Then
            Chk_Array(sw) = "Line2D"
            sw = sw + 1
        End If

        If Chk_6.CheckState = 1 Or Chk_1.CheckState = 1 Then
            Chk_Array(sw) = "Curve"
            sw = sw + 1
        End If

        If Chk_7.CheckState = 1 Or Chk_1.CheckState = 1 Then
            Chk_Array(sw) = "Plane"
            sw = sw + 1
        End If

        Chk_Array(sw) = "Length"
        sw = sw + 1
    End Function


    Private Sub Parameter_Create_Click(sender As Object, e As EventArgs) Handles Parameter_Create.Click
        ProgressBar1.Value = 0
        '--------------------------------------------------------------------------------------------------- 파일이 없을때...
        If Select_List.Items.Count = 0 Then
            aaa = MsgBox("File을 선택하세요.", vbInformation, "ISPark_Automation")
            Exit Sub
        End If
        '---------------------------------------------------------------------------------------------------  Parameter Name Text 값이 없을때...
        If Parameter_Name_Text.Text = "" Then
            aaa = MsgBox("Parameter Name 를 입력하세요", vbInformation, "ISPark_Automation")
            Exit Sub
        End If
        '--------------------------------------------------------------------------------------------------- Parameter Value Text 값이 없을때...
        If Parameter_Value_Text.Text = "" Then
            If Paramete_Create_Option.Checked = True Then
                aaa = MsgBox("Parameter Value 를 입력하세요", vbInformation, "ISPark_Automation")
                Exit Sub
            End If
        End If
        Message_Label.Text = "Parameter 구성 중..."
        ProgressBar1.Value = 10
        Call CATIA_BASIC02()
        '--------------------------------------------------------------------------------------------------- 파일이 수량이 1개일때...
        If file_number = 1 Then
            One_folder_path = Split(Folder_Name, Origin_File_Name)
            For i = 1 To 200
                Test_One_Origin_File_Name = Mid(Origin_File_Name, 1, i)
                If Test_One_Origin_File_Name = Origin_File_Name Then
                    Exit For
                End If
            Next
            One_Origin_File_Name = Mid(Origin_File_Name, 1, i - 11)
            '--------------------------------------------------------------------------------------------------- 읽기 쓰기 전용으로 변경
            fso = CreateObject("Scripting.FileSystemObject")
            file_Attributes_value = fso.GetFile(Folder_Name)
            file_Attributes_value.Attributes = 32
            ProgressBar1.Value = 30
            '--------------------------------------------------------------------------------------------------- File Open

            Call CATIA_BASIC02()
            partDocument = CATIA.Documents.Open(Folder_Name)
            selection = CATIA.ActiveDocument.Selection
            'partDocument = Documents.Open(Folder_Name)
            'partDocument.Visible = True

            Active_documents = CATIA.ActiveDocument
            parameters1 = Active_documents.Part.Parameters

            '--------------------------------------------------------------------------------------------------- Parameter 생성
            selection.search("(Name=" & Parameter_Name_Text.Text & "& CATKnowledgeSearch.InternalParameter),all")
            'selection.search("(Name=Parameter_Name_Text & CATKnowledgeSearch.InternalParameter),all")
            ProgressBar1.Value = 50
            '--------------------------------------------------------------------------------------------------- Search Parameter 값이 없을때만 적용...
            If selection.Count = 0 Then
                '--------------------------------------------------------------------------------------------------- Parameter 생성 Option 일때...
                If Paramete_Create_Option.Checked = True Then
                    strParam1 = parameters1.CreateString(Parameter_Name_Text.Text, "")
                    strParam1.Value = Parameter_Value_Text.Text
                    CATIA.ActiveDocument.Save()

                End If
            Else
                '--------------------------------------------------------------------------------------------------- Parameter 제거 Option 일때...
                If Paramete_Delete_Option.Checked = True Then
                    'selection.Add(selection.Item(1).Value) 6.0에서 사용했던 코드 삭제
                    selection.Delete
                    CATIA.ActiveDocument.Save()
                    ProgressBar1.Value = 70
                Else
                    selection.Clear
                End If
            End If
            CATIA.ActiveDocument.Close()
            '--------------------------------------------------------------------------------------------------- 파일 속성 변경 (읽기 전용)
            'Delay(0.5)
            file_Attributes_value.Attributes = 33
            ProgressBar1.Value = 90
        Else
            For i = 0 To file_number - 2
                ProgressBar1.Value = 10 + ((80 / file_number - 1) * (i + 1))
                File_Name(i) = Select_List.Items.Item(1)
                '--------------------------------------------------------------------------------------------------- 읽기 쓰기 전용으로 변경
                fso = CreateObject("Scripting.FileSystemObject")
                file_Attributes_value = fso.GetFile(Folder_Name & "\" & File_Name(i) & ".CATPart")
                file_Attributes_value.Attributes = 32
                '--------------------------------------------------------------------------------------------------- File Open
                partDocument = Documents.Open(Folder_Name & "\" & File_Name(i) & ".CATPart")
                Call CATIA_BASIC02()
                Active_documents = CATIA.ActiveDocument
                parameters1 = Active_documents.Part.Parameters
                '--------------------------------------------------------------------------------------------------- Parameter 생성
                selection.search("(Name=" & Parameter_Name_Text.Text & "& CATKnowledgeSearch.InternalParameter),all")
                '--------------------------------------------------------------------------------------------------- Search Parameter 값이 없을때만 적용...
                If selection.Count = 0 Then
                    '--------------------------------------------------------------------------------------------------- Parameter 생성 Option 일때...
                    If Paramete_Create_Option.Checked = True Then
                        strParam1 = parameters1.CreateString(Parameter_Name_Text.Text, "")
                        strParam1.Value = Parameter_Value_Text.Text
                        CATIA.ActiveDocument.Save()
                    End If
                Else
                    '--------------------------------------------------------------------------------------------------- Parameter 제거 Option 일때...
                    If Paramete_Delete_Option.Checked = True Then
                        selection.Add(selection.Item(1).Value)
                        selection.Delete
                        CATIA.ActiveDocument.Save()
                    End If
                    selection.Clear
                End If
                CATIA.ActiveDocument.Close()
                '--------------------------------------------------------------------------------------------------- 파일 속성 변경 (읽기 전용)
                'Delay(0.5)
                file_Attributes_value.Attributes = 33
            Next i
        End If
        Message_Label.Text = "Parameter 구성 완료"
        ProgressBar1.Value = 100
    End Sub

    Dim Coupling_Repeat_Folder, Chk_Array, Read_Only
    Public Search_Text_1, Search_Text_2
    Public StrFolder


    Private Sub File_Select_Button_Click(sender As Object, e As EventArgs) Handles File_Select_Button.Click
        '--------------------------------------------------------------------------------------------------- File Select
        CommonDialog1.FileName = ""
        CommonDialog1.Filter = "CATIAFile|*.CATPart"
        'CommonDialog1MaxFileSize = 20000
        'CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
        CommonDialog1.ShowDialog()
        '--------------------------------------------------------------------------------------------------- File 가져오기
        Name_Text.Text = Split(CommonDialog1.FileName, vbNullChar)(0)
        Select_List.Items.Clear()
        '--------------------------------------------------------------------------------------------------- File 선택이 없을때..
        If Name_Text.Text = "" Then
            aaa = MsgBox("File을 선택하세요.", vbInformation, "ISPark_Automation")
            Exit Sub
        End If
        '--------------------------------------------------------------------------------------------------- 파일 이름 Split
        Folder_Name = Split(CommonDialog1.FileName, vbNullChar)(0)
        '--------------------------------------------------------------------------------------------------- 폴더 Path 입력
        Data_Path_Text.Tag = Folder_Name
        '--------------------------------------------------------------------------------------------------- File Name 기입
        For file_number = 1 To UBound(Split(CommonDialog1.FileName, vbNullChar))
            Split_Name = Split(CommonDialog1.FileName, vbNullChar)(file_number)
            '--------------------------------------------------------------------------------------------------- File Name 비교
            For i = 1 To 200
                Test_Split_Name = Mid(Split_Name, 1, i)
                If Test_Split_Name = Split_Name Then
                    Exit For
                End If
            Next
            '--------------------------------------------------------------------------------------------------- List에 파일 추가
            Origin_File_Name = Mid(Test_Split_Name, 1, i - 8)
            Select_List.Items.Add(Origin_File_Name)
        Next
        '--------------------------------------------------------------------------------------------------- file_number가 1일때... File Name 비교
        If file_number = 1 Then
            ls_ss = Split(Folder_Name, "\")
            Origin_File_Name = ls_ss(UBound(ls_ss))
            Select_List.Items.Add(Origin_File_Name)
            '--------------------------------------------------------------------------------------------------- 폴더 Path 입력
            Data_Path_Text.Text = Split(CommonDialog1.FileName, "\" & Origin_File_Name)(0)
        End If
        Message_Label.Text = "Parameter 파일 선택 완료"
    End Sub



    Private Sub ZZZB_Parameter_Create_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Show()

        Call windows_ontop(HWND_TOPMOST)
        'Me.Top = 1080
        Me.CenterToScreen()
        'ME.Left = Screen.Width - 8000
        Me.Message_Label.Text = "Parameter 생성 파일 선택"

        Me.Chk_Option.CheckState = 0
        '    Me.Chk_Option.CheckState = 1
        Chk_1.CheckState = 1
        Process_Count = 1
        Call Chk_Option_Click()


        Me.List_Box.Items.Clear()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '   옵션 보기 - 체크박스
    '---------------------------------------------------------------------------------------------------
    Private Sub Chk_Option_Click()
        Call Btn_Clear_Click()

        If Chk_Option.CheckState = 0 Then
            Frm_Select.Visible = False
            SSTab2.Height = 1000

            Frm_State.Top = 4320

            Me.Height = 380
            Me.Width = 510

            If SSTab2.TabPages.Item(0).Visible = True Then
                Me.Height = 659
                Me.Width = 510
            End If

            Process_Count = 1
        ElseIf Chk_Option.CheckState = 1 Then
            Frm_Select.Visible = True

            SSTab2.Height = 1000

            'Frm_Select.Top = 4320
            'Frm_State.Top = 7560

            Me.Height = 659
            Me.Width = 510

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '   Clear 버튼 클릭.
    '---------------------------------------------------------------------------------------------------
    Private Sub Btn_Clear_Click()
        List_Box.Items.Clear()
        Lbl_Path.Text = "-"
        Lbl_State.Text = "-"
        ProgressBar3.Value = 0
    End Sub


End Class