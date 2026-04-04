# 한글 컨트롤(HwpCtrl) Action Table 전체 데이터

## 0. 범례 및 규칙 (Page 1 참조)
*   **Action ID 형식:**
    *   `빨간색 밑줄`: Dummy Action (실제 구현되지 않음, 내부 용도)
    *   `기울임 밑줄`: 한글 내부 기능을 모방하여 컨트롤에서 별도 구현한 액션
*   **ParameterSet ID 기호:**
    *   `-`: ParameterSet 없음.
    *   `+`: 추가 예정 (내부적 사용 중이나 외부 노출 미정).
    *   `*`: `HwpCtrl.Run` 불가능. 반드시 `CreateSet` -> `Execute` 과정을 거쳐야 함.

---

## 1. Action Table (A-Z 순서)

| Action ID | ParameterSet ID | Description | 비고 |
| :--- | :--- | :--- | :--- |
| **AddHanjaWord** | + | 한자단어 등록 | |
| **AllReplace** | FindReplace* | 모두 바꾸기 | |
| **AQcommandMerge** | UserQCommandFile* | 입력 자동 명령 파일 저장/로드 | |
| **AutoChangeHangul** | - | 낱자모 우선 | |
| **AutoChangeRun** | - | 동작 | |
| **AutoSpell Run** | - | 맞춤법 도우미 동작 On/Off | |
| **AutoSpellSelect1 ~ 16** | - | 맞춤법 도우미 선택 어휘로 변경 | |
| **Average** | Sum | 블록 평균 | |
| **BackwardFind** | FindReplace* | 뒤로 찾기 | |
| **Bookmark** | BookMark | 책갈피 | |
| **BookmarkEditDialog** | | 북마크 편집 대화상자 호출 | |
| **BreakColDef** | - | 단 정의 삽입 | |
| **BreakColumn** | - | 단 나누기 | |
| **BreakLine** | - | line break (줄 바꿈) | |
| **BreakPage** | - | 쪽 나누기 | |
| **BreakPara** | - | 문단 나누기 | |
| **BreakSection** | - | 구역 나누기 | |
| **BulletDlg** | ParaShape | 글머리표 대화상자 | |
| **Cancel** | - | ESC (작업 취소) | |
| **CaptionPosBottom** | ShapeObject | 캡션 위치-아래 | |
| **CaptionPosLeftBottom** | ShapeObject | 캡션 위치-왼쪽 아래 | |
| **CaptionPosLeftCenter** | ShapeObject | 캡션 위치-왼쪽 가운데 | |
| **CaptionPosLeftTop** | ShapeObject | 캡션 위치-왼쪽 위 | |
| **CaptionPosRightBottom** | ShapeObject | 캡션 위치-오른쪽 아래 | |
| **CaptionPosRightCenter** | ShapeObject | 캡션 위치-오른쪽 가운데 | |
| **CaptionPosRightTop** | ShapeObject | 캡션 위치-오른쪽 위 | |
| **CaptionPosTop** | ShapeObject | 캡션 위치-위 | |
| **CaptureDialog** | - | 갈무리 끝 | |
| **CaptureHandler** | - | 갈무리 시작 | |
| **CellBorder** | CellBorderFill | 셀 테두리 | |
| **CellBorderFill** | CellBorderFill | 셀 테두리 | |
| **CellFill** | CellBorderFill | 셀 배경 | |
| **CellZoneBorder** | CellBorderFill | 셀 테두리(여러 셀에 걸쳐 적용) | 셀 선택 시만 동작 |
| **CellZoneBorderFill** | CellBorderFill | 셀 테두리(여러 셀에 걸쳐 적용) | 셀 선택 시만 동작 |
| **CellZoneFill** | CellBorderFill | 셀 배경(여러 셀에 걸쳐 적용) | 셀 선택 시만 동작 |
| **ChangeImageFileExtension** | SummaryInfo | 연결 그림 확장자 바꾸기 | |
| **ChangeObject** | ShapeObject | 개체 변경하기 | |
| **ChangeRome String** | + | 로마자변환 - 입력받은 스트링 변환 | |
| **ChangeRome User String** | + | 로마자 사용자 데이터 추가 | |
| **ChangeRome User** | + | 로마자 사용자 데이터 | |
| **ChangeRome** | + | 로마자변환 | |
| **CharShape** | CharShape | 글자 모양 | |
| **CharShapeBold** | - | 단축키: Alt+L(글자 진하게) | |
| **CharShapeCenterline** | - | 취소선(CenterLine) | |
| **CharShapeDialog** | CharShape | 글자 모양 대화상자(내부 구현용) | |
| **CharShapeDialogWithoutBorder** | CharShape | 글자 모양 대화상자([글자 테두리] 탭 제외) | |
| **CharShapeEmboss** | - | 양각 | |
| **CharShapeEngrave** | - | 음각 | |
| **CharShapeHeight** | - | 글자 크기(대화상자 Focus이동용) | |
| **CharShapeHeightDecrease** | - | 크기 작게 (ALT+SHIFT+R) | |
| **CharShapeHeightIncrease** | - | 크기 크게 (ALT+SHIFT+E) | |
| **CharShapeItalic** | - | 이탤릭 (ALT+SHIFT+I) | |
| **CharShapeLang** | - | 글꼴 언어(대화상자 Focus이동용) | |
| **CharShapeNextFaceName** | - | 다음 글꼴 (ALT+SHIFT+F) | |
| **CharShapeNormal** | - | 보통모양 (ALT+SHIFT+C) | |
| **CharShapeOutline** | - | 외곽선 | |
| **CharShapePrevFaceName** | - | 이전 글꼴 (ALT+SHIFT+G) | |
| **CharShapeShadow** | - | 그림자 | |
| **CharShapeSpacing** | - | 자간(대화상자 Focus이동용) | |
| **CharShapeSpacingDecrease** | - | 자간 좁게 (ALT+SHIFT+N) | |
| **CharShapeSpacingIncrease** | - | 자간 넓게 (ALT+SHIFT+W) | |
| **CharShapeSubscript** | - | 아래첨자 (ALT+SHIFT+S) | |
| **CharShapeSuperscript** | - | 위첨자 (ALT+SHIFT+P) | |
| **CharShapeSuperSubscript** | - | 위첨자 -> 아래첨자 -> 보통 반복 | |
| **CharShapeTextColorBlack** | - | 글자색 검정 | |
| **CharShapeTextColorBlue** | - | 글자색 파랑 | |
| **CharShapeTextColorBluish** | - | 글자색 청록 | |
| **CharShapeTextColorGreen** | - | 글자색 초록 | |
| **CharShapeTextColorRed** | - | 글자색 빨강 | |
| **CharShapeTextColorViolet** | - | 글자색 자주 | |
| **CharShapeTextColorWhite** | - | 글자색 흰색 | |
| **CharShapeTextColorYellow** | - | 글자색 노랑 | |
| **CharShapeTypeFace** | - | 글꼴 이름(대화상자 Focus이동용) | |
| **CharShapeUnderline** | - | 밑줄 (ALT+SHIFT+U) | |
| **CharShapeWidth** | - | 장평(대화상자 Focus이동용) | |
| **CharShapeWidthDecrease** | - | 장평 좁게 (ALT+SHIFT+J) | |
| **CharShapeWidthIncrease** | - | 장평 넓게 (ALT+SHIFT+K) | |
| **Close** | - | 현재 리스트 닫고 상위 리스트로 이동 | |
| **CloseEx** | - | Close 확장 액션(Shift+Esc) | |
| **Comment** | - | 숨은 설명 | |
| **CommentDelete** | - | 숨은 설명 지우기 | |
| **CommentModify** | - | 숨은 설명 고치기 | |
| **CompatibleDocument** | CompatibleDocument | 호환 문서 | |
| **ComposeChars** | ChCompose | 글자 겹침 | |
| **ConvertBrailleSetting** | BrailleConvert | 점자 설정 변환 | |
| **ConvertCase** | ConvertCase | 대소문자 바꾸기 | |
| **ConvertFullHalfWidth** | ConvertFullHalf | 전각 반각 바꾸기 | |
| **ConvertHiraGata** | ConvertHiraToGata | 일어 바꾸기 | |
| **ConvertJianFan** | ConvertJianFan | 간/번체 바꾸기 | Text 선택 시 동작 |
| **ConvertOptGugyulToHangul** | ConvertToHangul | 한글로 옵션 - 구결을 한글로 | |
| **ConvertOptHanjaToHangul** | ConvertToHangul | 한글로 옵션 - 漢字를 한글로 | |
| **ConvertOptHanjaToHanjaHangul** | ConvertToHangul | 한글로 옵션 - 漢字를 漢字(한글)로 | |
| **ConvertToBraille** | BrailleConvert | 점자 변환 | |
| **ConvertToBrailleSelected** | BrailleConvert | 선택된 부분 점자 변환 | |
| **ConvertToHangul** | ConvertToHangul | 한글로 | |
| **Copy** | - | 복사하기 | |
| **CopyPage** | | 쪽 복사하기 | |
| **Cut** | - | 오려두기 | |
| **Delete** | - | 삭제 | |
| **DeleteBack** | - | Backspace (뒤로 지우기) | |
| **DeleteCtrls** | DeleteCtrls | 조판 부호 지우기 | |
| **DeleteDocumentMasterPage** | - | 문서 전체 바탕쪽 삭제 | |
| **DeleteDutmal** | + | 덧말 지우기 | |
| **DeleteField** | - | 누름틀/메모지우기 | 틀만 지우고 내용 유지 |
| **DeleteFieldMemo** | - | 메모 지우기 | |
| **DeleteLine** | - | CTRL-Y (한줄 지우기) | |
| **DeleteLineEnd** | - | ALT-Y (커서부터 줄 끝까지 지우기) | |
| **DeletePage** | DeletePage | 쪽 지우기 | |
| **DeletePrivateInfoMark** | - | 개인 정보 감추기 정보 다시보기 | |
| **DeletePrivateInfoMarkAtCurrentPos** | - | 현재 캐럿 위치 정보 다시 보기 | |
| **DeleteSectionMasterPage** | - | 구역 바탕쪽 삭제 | |
| **DeleteWord** | - | 단어 지우기 (CTRL-T) | |
| **DeleteWordBack** | - | CTRL-BS (앞 단어 지우기) | |
| **DocFindEnd** | FindReplace* | 문서 찾기 종료 | |
| **DocFindInit** | FindReplace* | 문서 찾기 초기화 | |
| **DocFindNext** | DocFindInfo* | 문서 찾기 계속 | |
| **DocSummaryInfo** | SummaryInfo | 문서 정보 | |
| **DocumentInfo** | DocumentInfo* | 현재 문서에 대한 정보 | |
| **DocumentSecurity** | DocSecurity | 문서 보안 설정 | |
| **DrawObjCancelOneStep** | - | 다각형/곡선 그리는 중 이전 선 지우기 | |
| **DrawObjCreatorArc** | ShapeObject | 호 그리기 | |
| **DrawObjCreatorCanvas** | ShapeObject | 캔버스 그리기 | |
| **DrawObjCreatorCurve** | ShapeObject | 곡선 그리기 | |
| **DrawObjCreatorEllipse** | ShapeObject | 원 그리기 | |
| **DrawObjCreatorFreeDrawing** | ShapeObject | 펜 그리기 | |
| **DrawObjCreatorHorzTextBox** | ShapeObject | 가로 글상자 만들기 | |
| **DrawObjCreatorLine** | ShapeObject | 선 그리기 | |
| **DrawObjCreatorMultiArc** | ShapeObject | 반복해서 호 그리기 | |
| **DrawObjCreatorMultiCanvas** | ShapeObject | 반복해서 캔버스 그리기 | |
| **DrawObjCreatorMultiCurve** | ShapeObject | 반복해서 곡선 그리기 | |
| **DrawObjCreatorMultiEllipse** | ShapeObject | 반복해서 원 그리기 | |
| **DrawObjCreatorMultiFreeDrawing** | ShapeObject | 반복해서 펜 그리기 | |
| **DrawObjCreatorMultiLine** | ShapeObject | 반복해서 선 그리기 | |
| **DrawObjCreatorMultiPolygon** | ShapeObject | 반복해서 다각형 그리기 | |
| **DrawObjCreatorMultiRectangle** | ShapeObject | 반복해서 사각형 그리기 | |
| **DrawObjCreatorMultiTextBox** | ShapeObject | 반복해서 글상자 그리기 | |
| **DrawObjCreatorObject** | ShapeObject | 그리기 개체 | |
| **DrawObjCreatorPolygon** | ShapeObject | 다각형 그리기 | |
| **DrawObjCreatorRectangle** | ShapeObject | 사각형 그리기 | |
| **DrawObjCreatorTextBox** | ShapeObject | 글상자 | |
| **DrawObjCreatorVertTextBox** | ShapeObject | 세로 글상자 만들기 | |
| **DrawObjEditDetail** | - | 그리기 개체 편집 | |
| **DrawObjOpenClosePolygon** | - | 다각형 열기/닫기 | |
| **DrawObjTemplateLoad** | ShapeObject | 그리기 마당에서 불러오기 | |
| **DrawObjTemplateSave** | - | 그리기 마당에 등록 | |
| **DrawShapeObjShadow** | ShapeObject | 개체 그림자 만들기/지우기 | 개체 선택 시만 동작 |
| **DropCap** | DropCap | 문단 첫 글자 장식 | |
| **DutmalChars** | Dutmal | 덧말 넣기 | |
| **EditFieldMemo** | - | 메모 내용 편집 | |
| **EditParaDown** | | 문단 아래로 옮기기 | |
| **EditParaUp** | | 문단 위로 옮기기 | |
| **EndnoteEndOfDocument** | SecDef | 미주-문서의 끝 | |
| **EndnoteEndOfSection** | SecDef | 미주-구역의 끝 | |
| **EndnoteToFootnote** | ExchangeFootnoteEndNote | 모든 미주를 각주로 | |
| **EquationCreate** | EqEdit | 수식 만들기 | |
| **EquationModify** | EqEdit | 수식 편집하기 | |
| **EquationPropertyDialog** | ShapeObject | 수식 개체 속성 고치기 | |
| **Erase** | - | 지우기 | |
| **ExchangeFootnoteEndnote** | ExchangeFootnoteEndNote | 각주/미주 변환 | |
| **ExecReplace** | FindReplace* | 바꾸기(실행) | |
| **ExtractImagesFromDoc** | SummaryInfo | 삽입 그림을 연결 그림으로 추출 | |
| **FileClose** | - | 문서 닫기 | |
| **FileNew** | - | 새 문서 | |
| **FileOpen** | - | 파일 열기 | |
| **FileOpenMRU** | - | 최근 작업 문서 열기 | |
| **FilePassword** | Password | 문서 암호 설정 | |
| **FilePasswordChange** | Password | 문서 암호 변경 및 해제 | |
| **FilePreview** | - | 미리 보기 | |
| **FileQuit** | - | 종료 | |
| **FileRWPasswordChange** | Password | 문서 열기/쓰기 암호 설정 | |
| **FileRWPasswordNew** | Password | 문서 열기/쓰기 암호 설정 | |
| **FileSave** | - | 파일 저장 | |
| **FileSaveAs** | - | 다른 이름으로 저장 | |
| **FileSaveAsImage** | Print | 그림 포맷으로 저장하기 | |
| **FileSaveAsImageOption** | Print | 그림 저장 옵션 설정 | |
| **FileSaveOptionDlg** | | 저장 옵션 대화상자 | |
| **FileSetSecurity** | FileSetSecurity* | 배포용 문서 만들기 | |
| **FileTemplate** | FileOpen | 문서마당 | |
| **FillColorShadeDec** | - | 면 색 음영 비율 감소 | |
| **FillColorShadeInc** | - | 면 색 음영 비율 증가 | |
| **FindAll** | FindReplace* | 모두 찾기 | |
| **FindDlg** | FindReplace | 찾기 | |
| **FindForeBackBookmark** | - | 앞뒤로 찾아가기 : 책갈피 | |
| **FindForeBackCtrl** | - | 앞뒤로 찾아가기 : 조판 부호 | |
| **FindForeBackFind** | - | 앞뒤로 찾아가기 : 찾기 | |
| **FindForeBackLine** | - | 앞뒤로 찾아가기 : 줄 | |
| **FindForeBackPage** | - | 앞뒤로 찾아가기 : 페이지 | |
| **FindForeBackSection** | - | 앞뒤로 찾아가기 : 구역 | |
| **FindForeBackStyle** | - | 앞뒤로 찾아가기 : 스타일 | |
| **FindOption** | FindReplace | 찾기 옵션 | 한/글 2024 부터 |
| **FootnoteBeneathText** | SecDef | 각주-본문 아래 | |
| **FootnoteBottomOfEachColumn** | SecDef | 다단 각주-각 단 아래 | |
| **FootnoteBottomOfMultiColumn** | SecDef | 다단 각주-전 단 | |
| **FootnoteBottomOfRightColumn** | SecDef | 다단 각주-오른쪽 단 아래 | |
| **FootnoteNoBeneathText** | SecDef | 각주-꼬리말 바로 위 | |
| **FootnoteOption** | SecDef | 각주/미주 모양 | |
| **FootnoteToEndnote** | ExchangeFootnoteEndNote | 모든 각주를 미주로 | |
| **FormDesignMode** | - | 디자인 모드 변경 | |
| **FormObjCreatorCheckButton** | - | Check버튼 넣기 | |
| **FormObjCreatorComboBox** | - | ComboBox넣기 | |
| **FormObjCreatorEdit** | - | Edit넣기 | |
| **FormObjCreatorListBox** | - | ListBox넣기 | |
| **FormObjCreatorPushButton** | - | Push버튼 넣기 | |
| **FormObjCreatorRadioButton** | - | Radio버튼 넣기 | |
| **FormObjCreatorScrollBar** | - | ScrollBar넣기 | |
| **ForwardFind** | FindReplace* | 앞으로 찾기 | |
| **FrameStatusBar** | - | 상태바 보이기/숨기기 | |
| **FtpDownload** | FtpDownload | FTP서버에서 파일 다운로드 및 오픈 | |
| **FtpUpload** | FtpUpload | 웹 서버로 올리기 | |
| **GetDefaultBullet** | ParaShape* | 글머리표 디폴트 값 가져오기 | |
| **GetDefaultParaNumber** | ParaShape* | 문단번호 디폴트 값 가져오기 | |
| **GetDocFilters** | DocFilters | 유틸리티 액션 | |
| **GetRome String** | ChangeRome* | 로마자 스펠링 얻기 | `Execute()` 필수 |
| **GetSectionApplyString** | SectionApply | 유틸리티 액션 | |
| **GetSectionApplyTo** | SectionApply | 유틸리티 액션 | |
| **GetVersionItemInfo** | VersionInfo | 저장된 버전/비교 Item 정보 얻기 | |
| **Goto** | GotoE | 찾아가기 | |
| **GotoStyle** | GotoE | 스타일 찾아가기 | |
| **HanThDIC** | - | 유의어 사전 | |
| **HeaderFooter** | HeaderFooter | 머리말/꼬리말 | |
| **HeaderFooterDelete** | - | 머리말 지우기 | |
| **HeaderFooterInsField** | HeaderFooter | 코드 넣기 | |
| **HeaderFooterModify** | - | 머리말/꼬리말 고치기 | |
| **HeaderFooterToNext** | - | 이후 머리말 이동 | |
| **HeaderFooterToPrev** | - | 이전 머리말 이동 | |
| **HiddenCredits** | - | 인터넷 정보 | |
| **HideTitle** | - | 차례 숨기기 | |
| **Him Config** | - | 입력기 언어별 환경설정 | |
| **HimKbdChange** | - | 키보드 바꾸기 | |
| **HwpCtrlEquationCreate97** | - | 수식 만들기(한글97버전) | |
| **HwpCtrlFileNew** | - | 새문서 | 한글 컨트롤 전용 |
| **HwpCtrlFileOpen** | - | 파일 열기 | 한글 컨트롤 전용 |
| **HwpCtrlFileSave** | - | 파일 저장 | 한글 컨트롤 전용 |
| **HwpCtrlFileSaveAs** | - | 다른 이름으로 저장 | 한글 컨트롤 전용 |
| **HwpCtrlFileSaveAsAutoBlock** | - | 블록 저장 | 한글 컨트롤 전용 |
| **HwpCtrlFindDlg** | - | 찾기 대화상자 | |
| **HwpCtrlReplaceDlg** | - | 바꾸기 대화상자 | |
| **HwpDic** | - | 한컴 사전 | |
| **Hyperlink** | HyperLink | 하이퍼링크 액션(Insert/Modify) | |
| **HyperlinkBackward** | - | 하이퍼링크 뒤로 | |
| **HyperlinkForward** | - | 하이퍼링크 앞으로 | |
| **HyperlinkJump** | HyperlinkJump | 하이퍼링크 이동 | |
| **Idiom** | Idiom | 상용구 | |
| **ImageFindPath** | - | 그림 경로 찾기 | |
| **ImportCharactersFromPicture** | | 그림에서 글자 가져오기 | |
| **IndexMark** | IndexMark | 찾아보기 표시 | |
| **IndexMarkModify** | IndexMark | 찾아보기 표시 고치기 | |
| **InputCodeChange** | - | 문자/코드 상호 변환 | 0x0020~0x10FFFF |
| **InputCodeTable** | CodeTable | 문자표 | |
| **InputDateStyle** | InputDateStyle | 날짜/시간 형식 지정하여 넣기 | |
| **InputHanja** | - | 한자 변환 | |
| **InputHanjaBusu** | - | 부수로 입력 | |
| **InputHanjaMean** | - | 새김으로 입력 | |
| **InputPersonsNameHanja** | InputHanja | 인명한자 변환 | |
| **InsertAutoNum** | - | 번호 다시 넣기 | |
| **InsertCCLMark** | HyperLink | CCL 넣기 | |
| **InsertChart** | OleCreation | 차트 만들기 | |
| **InsertConnectLine...** | ShapeObject | 각종 개체 연결선 그리기 | Arc, Straight, Stroke 등 |
| **InsertCpNo** | - | 상용구 코드 넣기(현재 쪽 번호) | |
| **InsertCpTpNo** | - | 상용구 코드 넣기(현재 쪽/전체 쪽) | |
| **InsertCrossReference** | ActionCrossRef | 상호 참조 만들기 | |
| **InsertDateCode** | - | 상용구 코드 넣기(만든 날짜) | |
| **InsertDocInfo** | - | 상용구 코드 넣기(작성자, 날짜 등) | |
| **InsertDocTitle** | InsertFieldTemplate | 상용구 코드 넣기(문서 제목) | |
| **InsertDocumentProperty** | InsertFieldTemplate | 상호 참조 넣기 | |
| **InsertEndnote** | - | 미주 입력 | |
| **InsertFieldCitation** | | 인용 삽입 | |
| **InsertFieldCtrl** | FieldCtrl | 필드 컨트롤(누름틀 등) 추가 | |
| **InsertFieldDateTime** | - | 날짜/시간 코드로 넣기 | |
| **InsertFieldFileName** | InsertFieldTemplate | 파일 이름 넣기 | |
| **InsertFieldMemo** | - | 메모 넣기 | |
| **InsertFieldRevisionChange** | - | 메모고침표 넣기 | |
| **InsertFieldTemplate** | InsertFieldTemplate | 문서마당 정보 삽입 | |
| **InsertFile** | InsertFile | 끼워 넣기 (파일 병합) | |
| **InsertFileName** | InsertFieldTemplate | 상용구 코드 넣기(파일 이름) | |
| **InsertFilePath** | InsertFieldTemplate | 상용구 코드 넣기(파일 경로 포함) | |
| **InsertFixedWidthSpace** | - | 고정폭 빈칸 삽입 | |
| **InsertFootnote** | - | 각주 입력 | |
| **InsertHyperlink** | HyperlinkJump | 하이퍼링크 만들기 | |
| **InsertIdiom** | Idiom | 상용구 등록 | |
| **InsertLastPrintDate** | - | 상용구 코드 넣기(마지막 인쇄 날짜) | |
| **InsertLastSaveBy** | - | 상용구 코드 넣기(마지막 저장한 사람) | |
| **InsertLastSaveDate** | - | 상용구 코드 넣기(마지막 저장 날짜) | |
| **InsertLine** | - | 선 넣기 | |
| **InsertLinkImageToDoc** | SummaryInfo | 연결 그림을 문서에 삽입 | |
| **InsertMovie** | ShapeObject | 동영상 파일 삽입 | |
| **InsertNonBreakingSpace** | - | 묶음 빈칸 삽입 | |
| **InsertPageNum** | - | 쪽 번호 넣기 | |
| **InsertRevision...** | RevisionDef | 각종 교정 부호 넣기 | 붙임, 뺌, 지움, 자리바꿈 등 |
| **InsertSoftHyphen** | - | 하이픈 삽입 | |
| **InsertSpace** | - | 공백 삽입 | |
| **InsertStringDateTime** | - | 날짜/시간 문자열로 넣기 | |
| **InsertTab** | - | 탭 삽입 | |
| **InsertText** | InsertText | 텍스트 삽입 | |
| **InsertTpNo** | - | 상용구 코드 넣기(전체 쪽수) | |
| **InsertUserName** | InsertFieldTemplate | 상용구 코드 넣기(만든 사람) | |
| **InsertVoice** | OleCreation | 음성 삽입 | |
| **Jajun** | - | 한자 자전 | |
| **LabelAdd** | - | 라벨 새 쪽 추가하기 | |
| **LabelTemplate** | - | 라벨 문서 만들기 | |
| **LeftShiftBlock** | | 블록 왼쪽 탭/공백 지우기 | |
| **LinkDocument** | LinkDocument | 문서 연결 | |
| **LinkTextBox** | - | 글상자 연결 | |
| **MacroDefine** | KeyMacro | 매크로 정의 | |
| **MacroPause** | - | 매크로 실행 일시 중지 | |
| **MacroPlay1 ~ 11** | - | 매크로 1~11 실행 | |
| **MacroRepeat** | - | 매크로 반복 실행 | |
| **MacroStop** | - | 매크로 실행 중지 | |
| **MailMergeField** | - | 메일 머지 필드 달기/고치기 | |
| **MailMergeGenerate** | MailMergeGenerate | 메일 머지 만들기 | |
| **MailMergeInsert** | FieldCtrl | 메일 머지 표시 달기 | |
| **MailMergeModify** | FieldCtrl | 메일 머지 고치기 | |
| **MakeAllVersionDiffs** | VersionInfo | 모든 버전비교 문서 만들기 | |
| **MakeContents** | MakeContents | 차례 만들기 | |
| **MakeIndex** | - | 찾아보기 만들기 | |
| **ManualChangeHangul** | - | 한영 수동 전환 | |
| **ManuScriptTemplate** | FileOpen | 원고지 쓰기 | |
| **MarkPenDelete** | | 형광펜 삭제 | |
| **MarkPenNext** | | 형광펜 이동(다음) | |
| **MarkPenPrev** | | 형광펜 이동(이전) | |
| **MarkPenShape** | MarkpenShape* | 형광펜 칠하기 | `Color` 값 설정 필수 |
| **MarkPrivateInfo** | PrivateInfoSecurity | 개인 정보 즉시 감추기 | 암호화 포함 |
| **MarkTitle** | - | 제목 차례 표시 | |
| **MasterPage** | MasterPage | 바탕쪽 | |
| **MasterPageDelete** | MasterPage* | 바탕쪽 삭제 | 편집모드에서만 |
| **MasterPageDuplicate** | - | 기존 바탕쪽과 겹침 | |
| **MasterPageEntry** | MasterPage | 바탕쪽 편집모드 진입 | |
| **MasterPageExcept** | - | 첫 쪽 제외 | |
| **MasterPageFront** | - | 바탕쪽 앞으로 보내기 | |
| **MasterPagePrevSection** | - | 앞 구역 바탕쪽 사용 | |
| **MasterPageToNext/Prev** | - | 이후/이전 바탕쪽 이동 | |
| **MasterPageTypeDlg** | MasterPage* | 바탕쪽 종류 다이얼로그 | |
| **MemoShape** | SecDef | 메모 모양 설정 | |
| **MemoToNext/Prev** | - | 다음/이전 메모 이동 | |
| **MessageBox** | + | 메시지 박스 출력 | |
| **ModifyBookmark** | BookMark | 책갈피 고치기 | |
| **ModifyComposeChars** | - | 글자 겹침 고치기 | |
| **ModifyCrossReference** | ActionCrossRef | 상호 참조 고치기 | |
| **ModifyCtrl** | - | 컨트롤 고치기 | |
| **ModifyDutmal** | - | 덧말 고치기 | |
| **ModifyField...** | InsertFieldTemplate | 각종 필드(날짜, 경로, 요약 등) 고치기 | |
| **ModifyFillProperty** | - | 채우기 속성 고치기 | |
| **ModifyLineProperty** | - | 선/테두리 속성 고치기 | |
| **ModifyRevision** | RevisionDef | 교정 부호 고치기 | 캐럿 위치 중요 |
| **ModifySecTextHorz/Vert** | TextVertical | 가로/세로 쓰기 설정 | |
| **ModifySection** | SecDef | 구역 설정 고치기 | |
| **ModifyShapeObject** | - | 개체 속성 고치기 | |
| **MoveColumnBegin/End** | - | 단의 시작/끝으로 이동 | |
| **MoveDocBegin/End** | - | 문서의 시작/끝으로 이동 | |
| **MoveDown/Up/Left/Right** | - | 캐럿 이동 (아래/위/왼쪽/오른쪽) | |
| **MoveLineBegin/End** | - | 줄의 시작/끝으로 이동 | |
| **MoveListBegin/End** | - | 리스트의 시작/끝으로 이동 | |
| **MoveNext/PrevChar** | - | 한 글자 뒤/앞 이동 | |
| **MoveNext/PrevPara** | - | 다음/이전 문단 이동 | |
| **MoveNext/PrevWord** | - | 한 단어 뒤/앞 이동 | |
| **MovePageBegin/End** | - | 페이지 시작/끝 이동 | |
| **MoveScrollDown/Up** | - | 화면 스크롤 이동 | |
| **MoveSel...** | - | 선택 블록(Selection)을 포함한 이동 | Doc, Line, Para 등 |
| **MultiColumn** | ColDef | 다단 설정 | |
| **NewNumber** | AutoNum | 새 번호로 시작 | |
| **NoteDelete/Modify** | - | 주석 지우기/고치기 | |
| **NoteSuperscript** | SecDef | 주석 번호 작게(윗첨자) | |
| **OleCreateNew** | OleCreation | OLE 개체 삽입 | |
| **OutlineNumber** | SecDef | 개요번호 설정 | |
| **PageBorder** | SecDef | 쪽 테두리/배경 | |
| **PageHiding** | PageHiding | 쪽 감추기 | |
| **PageSetup** | SecDef | 편집 용지 설정 | |
| **ParagraphShape** | ParaShape | 문단 모양 | |
| **ParagraphShapeAlign...** | - | 문단 정렬 (Left, Right, Center, Justify 등) | |
| **ParaNumberBullet** | ParaShape | 문단번호/글머리표 수준 조정 | |
| **Paste** | - | 붙이기 | |
| **PasteSpecial** | - | 골라 붙이기 | |
| **PictureChange** | PictureChange | 그림 바꾸기 | |
| **PictureEffect1 ~ 8** | - | 그림 효과(그레이, 흑백, 워터마크, 밝기 등) | |
| **PictureInsertDialog** | - | 그림 넣기 대화상자 | API용 |
| **PictureNo...** | ShapeObject | 그림 효과 제거(그림자, 네온, 반사 등) | |
| **Preference** | Preference | 환경 설정 | |
| **Print** | Print | 인쇄 실행 | |
| **PrintToPDF** | Print | PDF로 인쇄/저장 | |
| **QuickCorrect** | - | 빠른 교정 액션 | |
| **Redo** | - | 다시 실행 | |
| **ReplaceDlg** | FindReplace | 찾아 바꾸기 대화상자 | |
| **SaveBlockAction** | FileSaveBlock | 블록 저장하기 | |
| **SelectAll** | - | 모두 선택 | |
| **ShapeObjAlign...** | - | 개체 정렬 (Top, Bottom, Left, Right, Center 등) | |
| **ShapeObjGroup/Ungroup** | - | 개체 묶기/풀기 | |
| **Sort** | Sort | 소트 (정렬) | |
| **TableCreate** | TableCreation | 표 만들기 | |
| **TableDeleteRow/Column** | TableDeleteLine | 줄/칸 지우기 | |
| **TableMergeCell** | - | 셀 합치기 | |
| **TableSplitCell** | TableSplitCell | 셀 나누기 | |
| **Undo** | - | 되살리기 | |
| **ViewZoom** | ViewProperties | 화면 확대/축소 | |

---

---

# 한글 ParameterSet ID 구조 및 주요 Item 정리

## 1. 검색 및 탐색 계열 (Find & Replace)
가장 빈번하게 사용되며, 하위 세트를 포함하는 복합 구조입니다.

| ParameterSet ID | 주요 Item (Property) | 관련 Action |
| :--- | :--- | :--- |
| **FindReplace** | `FindString`(찾을 문자열), `ReplaceString`(바꿀 문자열), `IgnoreReplace`(대소문자 무시), `Forward`(정방향), `MatchCase`, `WholeWordOnly`, `FindCharShape`(하위 세트) | AllReplace, BackwardFind, FindDlg, RepeatFind |
| **GotoE** | `HwpPos`(위치), `Set` | Goto, GotoStyle |
| **DocFindInfo** | `FindString`, `Direction` | DocFindNext |

## 2. 텍스트 및 서식 계열 (Text Formatting)
글자 모양과 문단 모양을 결정하는 세트입니다.

| ParameterSet ID | 주요 Item (Property) | 관련 Action |
| :--- | :--- | :--- |
| **CharShape** | `Height`(크기), `FaceName`(글꼴), `TextColor`(글자색), `Bold`, `Italic`, `UnderlineType`, `UnderlineColor`, `Spacing`(자간), `Ratio`(장평) | CharShape, CharShapeBold 등 |
| **ParaShape** | `LeftMargin`, `RightMargin`, `Indentation`(들여쓰기), `LineSpacing`(줄 간격), `Alignment`(정렬: 0:양쪽, 1:왼쪽, 2:오른쪽, 3:가운데) | ParagraphShape, ParaShapeDialog, BulletDlg |
| **Style** | `StyleName`, `StyleType` | Style, StyleAdd, StyleEdit |

## 3. 표(Table) 조작 계열 (Table Operations)
표 생성 및 셀 테두리/배경을 조작합니다.

| ParameterSet ID | 주요 Item (Property) | 관련 Action |
| :--- | :--- | :--- |
| **TableCreation** | `Rows`(줄 수), `Cols`(칸 수), `WidthType`, `HeightType`, `CreateItemSet` | TableCreate |
| **CellBorderFill** | `BorderTypeLeft/Right/Top/Bottom`, `BorderWidth`, `FillColor`(배경색), `HatchColor`, `HatchStyle` | CellBorder, CellFill, TableCellShadeInc |
| **TableSplitCell** | `Rows`(나눌 줄 수), `Cols`(나눌 칸 수) | TableSplitCell |
| **TableDeleteLine** | `Side`(줄/칸 삭제 방향) | TableDeleteRow, TableDeleteColumn |

## 4. 개체 및 그리기 계열 (Shapes & Objects)
글상자, 선, 그림 등 개체 속성입니다.

| ParameterSet ID | 주요 Item (Property) | 관련 Action |
| :--- | :--- | :--- |
| **ShapeObject** | `Width`, `Height`, `TreatAsChar`(글자처럼 취급), `TextWrap`(본문과의 배치), `OffsetX`, `OffsetY`, `Caption`(캡션 내용) | DrawObjCreator..., ChangeObject, TextBoxAlign... |
| **ShapeCopyPaste** | `ShapeType`, `CopyMode` | ShapeCopyPaste, ShapeObjectCopy |
| **PictureChange** | `Path`(그림 경로), `Ext`(확장자) | PictureChange |

## 5. 문서 관리 및 보안 계열 (Document & Security)
파일 저장, 암호, 버전 관리에 쓰입니다.

| ParameterSet ID | 주요 Item (Property) | 관련 Action |
| :--- | :--- | :--- |
| **FileSetSecurity** | `Password`(암호), `NoPrint`(인쇄제한: true/false), `NoCopy`(복사제한: true/false) | FileSetSecurity (배포용 문서) |
| **SecDef** | `TopMargin`, `BottomMargin`, `LeftMargin`, `RightMargin`, `Gutter`(제본 여백), `FootnoteShape`(각주 모양) | PageSetup, FootnoteOption, PageBorder |
| **SummaryInfo** | `Author`(지은이), `Title`(제목), `Subject`(주제), `Keywords`(키워드) | DocSummaryInfo, ChangeImageFileExtension |
| **VersionInfo** | `VersionMemo`, `VersionDate` | SaveHistoryItem, GetVersionItemInfo |

## 6. 기타 특수 기능 계열
| ParameterSet ID | 용도 | 관련 Action |
| :--- | :--- | :--- |
| **HeaderFooter** | 머리말/꼬리말의 위치 및 내용 | HeaderFooter, HeaderFooterModify |
| **HyperLink** | 연결 종류, URL 주소, 대상 책갈피 | Hyperlink, InsertCCLMark |
| **RevisionDef** | 교정 부호의 종류 및 내용 | InsertRevision..., ModifyRevision |
| **EqEdit** | 수식 스크립트 내용 | EquationCreate, EquationModify |

---
