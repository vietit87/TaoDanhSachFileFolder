; ====================================================================
; DanhSachFileFolder.au3
; ====================================================================
#RequireAdmin

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=o.ico
#AutoIt3Wrapper_Res_Description=TaoDanhSachFileFolder_v1.2 ; tên hiển thị trong Task Manager
#AutoIt3Wrapper_Outfile=TaoDanhSachFileFolder_v1.2.exe ; tên file đầu ra (.exe) cho ứng dụng
#AutoIt3Wrapper_Res_Fileversion=1.2.0.0
#AutoIt3Wrapper_Res_Companyname=Copyright@lqviet_02.02.2026
#AutoIt3Wrapper_Res_Language=1066 ; Vietnamese
#AutoIt3Wrapper_Run_Obfuscator=y  ; Sử dụng bộ làm rối mã nguồn
#AutoIt3Wrapper_UseUpx=y   ;sử dụng công cụ UPX để nén file .exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <EditConstants.au3>

Global $g_sRoot = ""
Global $g_aList[0]
Global $g_iFolderCount = 0
Global $g_iFileCount = 0

; ================= GUI =================
GUICreate("Công cụ tạo danh sách tên tập tin và thư mục", 900, 680)

GUICtrlCreateLabel("Phát triển bởi Lê Quốc Việt", 10, 10, 400, 20)

Global $inpFolder = GUICtrlCreateInput("", 10, 40, 650, 25)
Global $btnBrowse = GUICtrlCreateButton("Duyệt...", 670, 40, 100, 25)

Global $lblSelected = GUICtrlCreateLabel("Thư mục đang chọn: chưa chọn thư mục nào", 10, 70, 850, 20)

GUICtrlCreateGroup("Chức năng", 10, 100, 870, 120)

; ===== Dòng 1 chức năng =====
Global $chkFolder = GUICtrlCreateCheckbox("Thư mục", 20, 125, 120, 20)
GUICtrlSetState($chkFolder, $GUI_CHECKED)

Global $chkFile = GUICtrlCreateCheckbox("Tập tin", 150, 125, 120, 20)
GUICtrlSetState($chkFile, $GUI_CHECKED)

Global $chkFolder1 = GUICtrlCreateCheckbox("Thư mục cấp 1", 300, 125, 150, 20)
Global $chkFolder2 = GUICtrlCreateCheckbox("Thư mục cấp 2", 470, 125, 150, 20)

Global $chkOnlyFile = GUICtrlCreateCheckbox("Chỉ tập tin", 640, 125, 110, 20)
Global $chkOnlyFolder = GUICtrlCreateCheckbox("Chỉ thư mục", 760, 125, 120, 20)

; ===== Dòng 2 =====
Global $chkExt = GUICtrlCreateCheckbox("Hiển thị đuôi tập tin", 20, 155, 180, 20)
GUICtrlSetState($chkExt, $GUI_CHECKED)

Global $chkFullPath = GUICtrlCreateCheckbox("Hiển thị đường dẫn đầy đủ", 220, 155, 220, 20)

Global $chkTxt = GUICtrlCreateCheckbox("Xuất txt", 470, 155, 100, 20)
GUICtrlSetState($chkTxt, $GUI_CHECKED)

Global $chkHtml = GUICtrlCreateCheckbox("Xuất html", 580, 155, 100, 20)

Global $chkExcel = GUICtrlCreateCheckbox("Xuất Excel", 690, 155, 120, 20)

Global $listResult = GUICtrlCreateEdit("", 10, 230, 870, 350, BitOR($ES_AUTOVSCROLL, $WS_VSCROLL))

Global $btnShow = GUICtrlCreateButton("Hiển thị", 330, 600, 100, 30)
Global $btnExport = GUICtrlCreateButton("Xuất KQ", 450, 600, 100, 30)
Global $btnExit = GUICtrlCreateButton("Thoát", 570, 600, 100, 30)

Global $lblResult = GUICtrlCreateLabel("", 10, 640, 870, 20)
GUICtrlSetColor($lblResult, 0x008000)
GUICtrlSetFont($lblResult, 10, 400, 2)

GUISetState()

; ================= FUNCTIONS =================

Func BrowseFolder()
    Local $s = FileSelectFolder("Chọn thư mục", "")
    If $s <> "" Then
        $g_sRoot = $s
        GUICtrlSetData($inpFolder, $s)
        Local $name = StringTrimLeft($s, StringInStr($s, "\", 0, -1))
        GUICtrlSetData($lblSelected, "Thư mục đang chọn: " & $name)
        BuildList()
    EndIf
EndFunc

; ================= BUILD =================

Func BuildList()

    If $g_sRoot = "" Then
        MsgBox(48, "Thông báo", "Vui lòng chọn thư mục trước.")
        Return
    EndIf

    ReDim $g_aList[0]
    $g_iFolderCount = 0
    $g_iFileCount = 0

	; ===== THƯ MỤC (chỉ thư mục con của thư mục gốc) =====
	If GUICtrlRead($chkFolder) = $GUI_CHECKED Then
		_GetOnlyFolders()
	EndIf

	; ===== THƯ MỤC CẤP 1 =====
	If GUICtrlRead($chkFolder1) = $GUI_CHECKED Then
		_GetLevel1Folders()
	EndIf

	; ===== THƯ MỤC CẤP 2 =====
	If GUICtrlRead($chkFolder2) = $GUI_CHECKED Then
		_GetLevel2Folders()
	EndIf

	; ===== FILE =====
	If GUICtrlRead($chkFile) = $GUI_CHECKED Then
		_GetFilesByOption()
	EndIf

    If UBound($g_aList) > 0 Then _ArraySort($g_aList)

    GUICtrlSetData($listResult, _ArrayToString($g_aList, @CRLF))

    GUICtrlSetData($lblResult, _
        "Kết quả: đã lập danh sách " & UBound($g_aList) & _
        " đối tượng, gồm " & $g_iFolderCount & _
        " thư mục và " & $g_iFileCount & " tập tin")
EndFunc

Func _GetOnlyFolders()

    Local $aFolders = _FileListToArray($g_sRoot, "*", $FLTA_FOLDERS)
    If @error Then Return

    For $i = 1 To $aFolders[0]
        _AddItem($g_sRoot & "\" & $aFolders[$i], True)
    Next

EndFunc

Func _AllowDisplay($isFolder)

    Local $onlyFile = GUICtrlRead($chkOnlyFile) = $GUI_CHECKED
    Local $onlyFolder = GUICtrlRead($chkOnlyFolder) = $GUI_CHECKED

    ; nếu không tick gì → cho phép tất cả
    If Not $onlyFile And Not $onlyFolder Then Return True

    ; nếu tick cả hai → cho phép tất cả
    If $onlyFile And $onlyFolder Then Return True

    ; nếu chỉ tập tin
    If $onlyFile And Not $isFolder Then Return True

    ; nếu chỉ thư mục
    If $onlyFolder And $isFolder Then Return True

    Return False
EndFunc

Func _AddSubFolders()
    Local $aFolders = _FileListToArray($g_sRoot, "*", $FLTA_FOLDERS)
    If @error Then Return

    For $i = 1 To $aFolders[0]
        Local $sub = $g_sRoot & "\" & $aFolders[$i]

        _AddFolder($sub)

        If GUICtrlRead($chkFile) = 1 Then
            _AddFiles($sub)
        EndIf
    Next
EndFunc

Func _AddFolder($fullPath)
    $g_iFolderCount += 1
    _AddItem($fullPath, True)
EndFunc

Func _AddFiles($path)
    Local $aFiles = _FileListToArray($path, "*", $FLTA_FILES)
    If @error Then Return

    For $i = 1 To $aFiles[0]
        $g_iFileCount += 1
        _AddItem($path & "\" & $aFiles[$i], False)
    Next
EndFunc

; ================= LEVEL 1 =================

Func _GetLevel1Folders()

    Local $aFolders = _FileListToArray($g_sRoot, "*", $FLTA_FOLDERS)
    If @error Then Return

    For $i = 1 To $aFolders[0]
        _AddItem($g_sRoot & "\" & $aFolders[$i], True)
    Next

EndFunc


; ================= LEVEL 2 =================

Func _GetLevel2Folders()

    Local $aFolders = _FileListToArray($g_sRoot, "*", $FLTA_FOLDERS)
    If @error Then Return

    For $i = 1 To $aFolders[0]

        Local $level1Path = $g_sRoot & "\" & $aFolders[$i]
        Local $aLevel2 = _FileListToArray($level1Path, "*", $FLTA_FOLDERS)

        If Not @error Then
            For $j = 1 To $aLevel2[0]
                _AddItem($level1Path & "\" & $aLevel2[$j], True)
            Next
        EndIf

    Next

EndFunc


; ================= FILE =================

Func _GetFilesByOption()

    ; file trong thư mục gốc
    _GetFilesInFolder($g_sRoot)

    ; nếu tick cấp 1 thì lấy file trong cấp 1
    If GUICtrlRead($chkFolder1) = 1 Or GUICtrlRead($chkFolder2) = 1 Then

        Local $aFolders = _FileListToArray($g_sRoot, "*", $FLTA_FOLDERS)
        If @error Then Return

        For $i = 1 To $aFolders[0]

            Local $level1Path = $g_sRoot & "\" & $aFolders[$i]
            _GetFilesInFolder($level1Path)

            ; nếu tick cấp 2 thì lấy file trong cấp 2
            If GUICtrlRead($chkFolder2) = 1 Then

                Local $aLevel2 = _FileListToArray($level1Path, "*", $FLTA_FOLDERS)

                If Not @error Then
                    For $j = 1 To $aLevel2[0]
                        _GetFilesInFolder($level1Path & "\" & $aLevel2[$j])
                    Next
                EndIf

            EndIf

        Next

    EndIf

EndFunc


Func _GetFilesInFolder($path)

    Local $aFiles = _FileListToArray($path, "*", $FLTA_FILES)
    If @error Then Return

    For $i = 1 To $aFiles[0]
        _AddItem($path & "\" & $aFiles[$i], False)
    Next

EndFunc

; ================= ADD ITEM =================

Func _AddItem($fullPath, $isFolder)

    ; ===== BỘ LỌC CHỈ TẬP TIN / CHỈ THƯ MỤC =====
    If Not _AllowDisplay($isFolder) Then Return

    Local $displayName

    ; Full path hay chỉ tên
    If GUICtrlRead($chkFullPath) = 1 Then
        $displayName = $fullPath
    Else
        $displayName = StringTrimLeft($fullPath, StringInStr($fullPath, "\", 0, -1))
    EndIf

    ; Nếu là file và KHÔNG tick hiển thị đuôi thì bỏ đuôi
    If Not $isFolder Then
        If GUICtrlRead($chkExt) <> $GUI_CHECKED Then
            $displayName = StringRegExpReplace($displayName, "\.[^.]*$", "")
        EndIf
    EndIf

	; thêm vào mảng
	ReDim $g_aList[UBound($g_aList) + 1]
	$g_aList[UBound($g_aList) - 1] = $displayName

	; ===== ĐẾM CHÍNH XÁC SAU LỌC =====
	If $isFolder Then
		$g_iFolderCount += 1
	Else
		$g_iFileCount += 1
	EndIf

EndFunc

Func _ExportToExcel($filePath)

    Local $aData = $g_aList
    If UBound($aData) = 0 Then Return

    Local $oExcel = ObjCreate("Excel.Application")
    If @error Then
        MsgBox(16, "Lỗi", "Không khởi tạo được Excel.")
        Return
    EndIf

    $oExcel.Visible = False
    Local $oBook = $oExcel.Workbooks.Add
    Local $oSheet = $oBook.Worksheets(1)

    For $i = 0 To UBound($aData) - 1
        $oSheet.Cells($i + 1, 1).Value = $aData[$i]
    Next

    $oBook.SaveAs($filePath)
    $oBook.Close
    $oExcel.Quit

    $oSheet = 0
    $oBook = 0
    $oExcel = 0

EndFunc

; ================= Export Result =================

Func ExportResult()

    If GUICtrlRead($chkTxt) = 0 And _
       GUICtrlRead($chkHtml) = 0 And _
       GUICtrlRead($chkExcel) = 0 Then
        MsgBox(48, "Thông báo", "Chọn định dạng xuất.")
        Return
    EndIf

    Local $content = GUICtrlRead($listResult)

    ; ===== TXT =====
    If GUICtrlRead($chkTxt) = 1 Then
        Local $f = FileSaveDialog("Lưu TXT", "", "Text (*.txt)", 2)
        If $f <> "" Then FileWrite($f, $content)
    EndIf

    ; ===== HTML =====
    If GUICtrlRead($chkHtml) = 1 Then
        Local $f = FileSaveDialog("Lưu HTML", "", "HTML (*.html)", 2)
        If $f <> "" Then
            FileWrite($f, "<html><body><pre>" & $content & "</pre></body></html>")
        EndIf
    EndIf

    ; ===== EXCEL =====
    If GUICtrlRead($chkExcel) = 1 Then
        Local $f = FileSaveDialog("Lưu Excel", "", "Excel (*.xlsx)", 2)
        If $f <> "" Then
            _ExportToExcel($f)
        EndIf
    EndIf

EndFunc

; ================= LOOP =================
While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE, $btnExit
            Exit
        Case $btnBrowse
            BrowseFolder()
        Case $btnShow
            BuildList()
        Case $btnExport
            ExportResult()
        Case $chkFile, $chkFolder, $chkFolder1, $chkFolder2, _
			$chkExt, $chkFullPath, $chkOnlyFile, $chkOnlyFolder
            BuildList()
    EndSwitch
WEnd
