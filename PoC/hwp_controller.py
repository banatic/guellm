"""
hwp_controller.py
HWP COM 자동화 레이어

【설계 원칙】
HWP COM은 STA(Single Threaded Apartment) 방식.
반드시 COM 객체를 생성한 스레드(메인 스레드)에서만 호출해야 합니다.
모든 public 메서드는 메인 스레드에서 직접 호출하세요.
"""
import os
import json
import time
from html.parser import HTMLParser


class HwpController:
    """한글(HWP) COM 자동화 컨트롤러 — 메인 스레드 전용"""

    def __init__(self):
        self.hwp = None
        self._connected = False
        self._cached_paragraphs: list[dict] = []

    # ──────────────────────────────────────────────
    # 연결 / 파일 관리
    # ──────────────────────────────────────────────

    def connect(self, visible: bool = True) -> bool:
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        self.hwp = win32.Dispatch("HWPFrame.HwpObject")
        self.hwp.XHwpWindows.Item(0).Visible = visible
        try:
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        self._connected = True
        return True

    def open_file(self, path: str) -> bool:
        self._ensure()
        abs_path = os.path.abspath(path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"파일 없음: {abs_path}")
        ext = abs_path.lower().rsplit(".", 1)[-1]
        if ext not in ("hwp", "hwpx"):
            raise ValueError("HWP/HWPX 파일만 지원합니다.")
        self._cached_paragraphs = []
        fmt = "HWPX" if ext == "hwpx" else "HWP"
        result = self.hwp.Open(abs_path, fmt, "forceopen:true")
        if not result:
            result = self.hwp.Open(abs_path, "", "forceopen:true")
        return bool(result)

    def save(self, save_path: str = None) -> str:
        self._ensure()
        if save_path:
            abs_path = os.path.abspath(save_path)
            try:
                act = self.hwp.CreateAction("FileSaveAs_S")
                ps = act.CreateSet()
                ps.SetItem("filename", abs_path)
                ps.SetItem("Format", "HWP")
                ps.SetItem("Attributes", 0)
                act.Execute(ps)
            except Exception:
                self.hwp.SaveAs(abs_path, "HWP", "")
            return abs_path
        else:
            try:
                act = self.hwp.CreateAction("FileSave_S")
                act.Execute(act.CreateSet())
            except Exception:
                self.hwp.Save()
            return self.hwp.Path

    def close(self, save: bool = False):
        try:
            if self._connected and self.hwp:
                if not save:
                    self.hwp.Quit()
        except Exception:
            pass
        finally:
            self._connected = False

    # ──────────────────────────────────────────────
    # 카테고리 1: Inspection (문서 구조 파악)
    # ──────────────────────────────────────────────

    def analyze_document_structure(self) -> dict:
        """
        문서의 전체 페이지 수, 표(Table) 개수, 누름틀(Field) 목록을 요약 반환합니다.
        LLM 작업 시작 시 항상 먼저 호출하세요.
        """
        self._ensure()
        result = {"pages": 0, "table_count": 0, "fields": [], "paragraph_count": 0}

        # 페이지 수
        try:
            result["pages"] = int(self.hwp.PageCount)
        except Exception:
            pass

        # 표 개수 + 단락 수 (HTML 파싱)
        html = self._get_html()
        if html:
            result["table_count"] = html.lower().count("<table")
            result["paragraph_count"] = html.lower().count("<p")

        # 필드 목록
        result["fields"] = self._get_field_names()

        return result

    def get_field_info(self) -> list[dict]:
        """
        문서 내 모든 누름틀(필드)의 이름과 현재 입력된 값을 반환합니다.
        """
        self._ensure()
        fields = []
        for name in self._get_field_names():
            try:
                value = self.hwp.GetFieldText(name) or ""
            except Exception:
                value = ""
            fields.append({"name": name, "value": value})
        return fields

    def get_table_schema(self, table_index: int) -> dict:
        """
        특정 표의 행/열 개수, 헤더, 각 셀 텍스트를 반환합니다.
        ActionTable: TableCreation(Rows, Cols), HTML 파싱으로 구조 추출.
        """
        self._ensure()
        html = self._get_html()
        if not html:
            return {"rows": 0, "cols": 0, "headers": [], "cells": []}

        tables = self._parse_all_tables(html)
        if table_index >= len(tables):
            return {"rows": 0, "cols": 0, "headers": [],
                    "cells": [], "error": f"표 {table_index}번 없음 (총 {len(tables)}개)"}

        table = tables[table_index]
        rows = len(table)
        cols = max((len(row) for row in table), default=0)
        return {
            "table_index": table_index,
            "rows": rows,
            "cols": cols,
            "headers": table[0] if table else [],
            "cells": table,
        }

    def get_all_tables_overview(self) -> list[dict]:
        """
        문서 내 모든 표의 인덱스, 행/열 수, 헤더(첫 행)를 한 번에 반환합니다.
        어떤 표를 수정해야 할지 파악할 때 get_table_schema를 반복 호출하는 대신 이 도구를 먼저 사용하세요.
        """
        self._ensure()
        html = self._get_html()
        if not html:
            return []
        tables = self._parse_all_tables(html)
        overview = []
        for idx, table in enumerate(tables):
            rows = len(table)
            cols = max((len(r) for r in table), default=0)
            headers = table[0] if table else []
            # 헤더가 너무 길면 축약
            short_headers = [h[:30] if len(h) > 30 else h for h in headers]
            overview.append({
                "table_index": idx,
                "rows": rows,
                "cols": cols,
                "headers": short_headers,
            })
        return overview

    def find_text_anchor(self, keyword: str) -> dict:
        """
        키워드 텍스트가 문서에 존재하는지 확인합니다.
        HTML 사전 확인으로 ForwardFind 호출 전에 존재 여부를 검증합니다.
        (ForwardFind가 텍스트를 못 찾으면 HWP 다이얼로그 블로킹 발생)
        """
        self._ensure()
        # HTML에서 먼저 확인 — 없으면 ForwardFind 호출 자체를 건너뜀
        if not self._text_exists(keyword, case_sensitive=False):
            return {"found": False, "keyword": keyword}
        try:
            act = self.hwp.CreateAction("ForwardFind")
            ps = act.CreateSet()
            ps.SetItem("FindString", keyword)
            ps.SetItem("IgnoreCase", True)
            ps.SetItem("Direction", 3)   # 문서 전체 탐색 — wrap-around 다이얼로그 방지
            ps.SetItem("FindReplace", 0)
            found = bool(act.Execute(ps))
            return {"found": found, "keyword": keyword}
        except Exception as e:
            return {"found": False, "keyword": keyword, "error": str(e)}

    # ──────────────────────────────────────────────
    # 카테고리 2: Smart Filling (데이터 입력)
    # ──────────────────────────────────────────────

    def fill_field_data(self, data_map: dict) -> str:
        """
        {'필드이름': '값'} 형태의 데이터를 누름틀에 채웁니다.
        ActionTable: FieldCtrl ParameterSet — hwp.SetFieldText() API 사용.
        """
        self._ensure()
        results = []
        for field_name, value in data_map.items():
            try:
                self.hwp.SetFieldText(field_name, str(value))
                results.append(f"✅ {field_name}={value}")
            except Exception as e:
                results.append(f"❌ {field_name}: {e}")
        return "\n".join(results)

    def replace_text_patterns(self, mapping: dict) -> str:
        """
        {{이름}}, [날짜] 등 템플릿 패턴을 실제 데이터로 일괄 치환합니다.
        ActionTable: AllReplace + FindReplace* ParameterSet
        """
        self._ensure()
        results = []
        for pattern, value in mapping.items():
            try:
                self._find_replace(pattern, str(value), case_sensitive=True)
                results.append(f"✅ '{pattern}' → '{value}'")
            except Exception as e:
                results.append(f"❌ '{pattern}': {e}")
        self._cached_paragraphs = []
        return "\n".join(results)

    def set_checkbox_state(self, label: str, is_checked: bool) -> str:
        """
        "label [ ]" 또는 "[ ] label" 형태의 체크박스 상태를 변경합니다.
        ActionTable: AllReplace 사용.
        """
        self._ensure()
        check_char = "V" if is_checked else " "
        undo_char = " " if is_checked else "V"
        changed = 0
        for sep in (" ", ""):
            # 앞에 레이블이 오는 경우
            for box in ("[ ]", f"[{undo_char}]"):
                old = f"{label}{sep}{box}"
                new = f"{label}{sep}[{check_char}]"
                try:
                    self._find_replace(old, new, case_sensitive=True)
                    changed += 1
                except Exception:
                    pass
            # 뒤에 레이블이 오는 경우
            for box in ("[ ]", f"[{undo_char}]"):
                old = f"{box}{sep}{label}"
                new = f"[{check_char}]{sep}{label}"
                try:
                    self._find_replace(old, new, case_sensitive=True)
                    changed += 1
                except Exception:
                    pass
        state_str = "체크" if is_checked else "해제"
        return f"체크박스 '{label}' {state_str} 처리 완료"

    def insert_image_box(self, anchor_text: str, image_path: str,
                         size_mode: str = "fit") -> str:
        """
        특정 텍스트 위치에 이미지를 삽입합니다.
        ActionTable: ShapeObject ParameterSet + hwp.InsertPicture()
        size_mode: "fit"=셀에 맞춤, "original"=원본 크기
        """
        self._ensure()
        abs_path = os.path.abspath(image_path)
        if not os.path.exists(abs_path):
            return f"❌ 이미지 파일 없음: {abs_path}"
        try:
            # anchor_text 위치로 커서 이동
            if anchor_text:
                act = self.hwp.CreateAction("ForwardFind")
                ps = act.CreateSet()
                ps.SetItem("FindString", anchor_text)
                ps.SetItem("IgnoreCase", True)
                act.Execute(ps)
            # sizetype: 0=원본, 1=셀크기에 맞춤, 2=비율 유지 확대
            sizetype = 1 if size_mode == "fit" else 0
            self.hwp.InsertPicture(abs_path, True, sizetype, False, False)
            return f"✅ 이미지 삽입: {os.path.basename(abs_path)}"
        except Exception as e:
            return f"❌ 이미지 삽입 실패: {e}"

    # ──────────────────────────────────────────────
    # 카테고리 3: Dynamic Table (표 동적 제어)
    # ──────────────────────────────────────────────

    def sync_table_rows(self, table_index: int, data_count: int,
                        header_rows: int = 1) -> str:
        """
        입력 데이터 개수에 맞춰 표의 데이터 행을 추가/삭제합니다.
        ActionTable: TableInsertLowerRow (Run 가능), TableDeleteRow (Execute 필요)
        """
        self._ensure()
        schema = self.get_table_schema(table_index)
        if schema["rows"] == 0:
            return f"❌ 표 {table_index}번을 찾을 수 없습니다."

        current_data_rows = schema["rows"] - header_rows
        diff = data_count - current_data_rows

        if diff == 0:
            return f"변경 없음 (현재 {current_data_rows}행)"

        # 커서를 표의 마지막 행에 위치시키기 위해 anchor 텍스트 탐색
        last_row = schema["cells"][-1] if schema["cells"] else []
        anchor = next((" ".join(c.split()) for c in last_row if c.strip()), None)
        if anchor and self._text_exists(anchor):
            try:
                act = self.hwp.CreateAction("ForwardFind")
                ps = act.CreateSet()
                ps.SetItem("FindString", anchor)
                ps.SetItem("IgnoreCase", True)
                ps.SetItem("Direction", 3)
                ps.SetItem("FindReplace", 0)
                act.Execute(ps)
            except Exception:
                pass

        if diff > 0:
            for _ in range(diff):
                self.hwp.Run("TableInsertLowerRow")
            return f"✅ 행 {diff}개 추가 (총 {data_count}행)"
        else:
            for _ in range(-diff):
                try:
                    act = self.hwp.CreateAction("TableDeleteRow")
                    ps = act.CreateSet()
                    act.Execute(ps)
                except Exception:
                    self.hwp.Run("TableDeleteRow")
            return f"✅ 행 {-diff}개 삭제 (총 {data_count}행)"

    def fill_table_data_matrix(self, table_index: int, start_row: int,
                                matrix: list, cell_delay: float = 0.3) -> str:
        """
        2차원 배열 데이터를 표의 특정 행부터 채웁니다.

        커서 탐색 + InsertText 방식 (AllReplace 방식 폐기):
        - AllReplace는 짧은 문자열의 다중 교체, _get_html 반복 호출로 인해 프리징 유발
        - 헤더 셀 텍스트로 표 내부로 진입한 뒤 TableRightCell로 셀 이동,
          SelectAll → InsertText로 셀 단위 정확 교체
        """
        self._ensure()
        schema = self.get_table_schema(table_index)
        if schema["rows"] == 0:
            return f"❌ 표 {table_index}번을 찾을 수 없습니다."

        cols = schema["cols"]
        filled, skipped = 0, []

        # ── 표 내부로 커서 진입 ──────────────────────────────────────────
        # 헤더 첫 번째 셀의 텍스트로 ForwardFind → 표 내부에 커서 위치
        header_cells = schema.get("headers", [])
        anchor = next((" ".join(c.split()) for c in header_cells if len(c.strip()) >= 2), None)
        if not anchor:
            return "❌ 표 헤더에서 anchor 텍스트를 찾을 수 없습니다."
        if not self._text_exists(anchor):
            return f"❌ anchor 텍스트를 문서에서 찾을 수 없습니다: {anchor!r}"

        self.hwp.Run("MoveDocBegin")
        act = self.hwp.CreateAction("ForwardFind")
        ps = act.CreateSet()
        ps.SetItem("FindString", anchor)
        ps.SetItem("IgnoreCase", False)
        ps.SetItem("Direction", 3)   # 3=문서 전체 탐색 — wrap-around 다이얼로그 방지
        ps.SetItem("FindReplace", 0) # 찾기 전용 (바꾸기 아님)
        if not act.Execute(ps):
            return f"❌ ForwardFind 실패: {anchor!r}"

        # ForwardFind 후 커서는 찾은 텍스트 끝에 위치 → 셀[0,0] 안에 있음
        # start_row행 첫 번째 열까지 TableRightCell로 이동
        tabs_needed = start_row * cols  # 헤더 이후 start_row행 첫 열
        for _ in range(tabs_needed):
            self.hwp.Run("TableRightCell")

        # ── 행/열 루프: 셀 이동 + 내용 교체 ────────────────────────────
        for r_offset, row_data in enumerate(matrix):
            for c_idx, new_text in enumerate(row_data):
                if c_idx >= cols:
                    skipped.append(f"행{start_row+r_offset} col{c_idx} 열 범위 초과")
                    continue
                try:
                    # 셀 내용 전체 선택 후 새 텍스트 삽입
                    self.hwp.Run("SelectAll")
                    ins = self.hwp.CreateAction("InsertText")
                    ips = ins.CreateSet()
                    ips.SetItem("Text", str(new_text))
                    ins.Execute(ips)
                    filled += 1
                    if cell_delay > 0:
                        time.sleep(cell_delay)
                except Exception as e:
                    skipped.append(f"셀[{start_row+r_offset},{c_idx}]: {e}")

                # 다음 셀로 이동 (마지막 셀이면 이동 안 함)
                if c_idx < len(row_data) - 1:
                    self.hwp.Run("TableRightCell")

            # 다음 행 첫 번째 열로 이동
            if r_offset < len(matrix) - 1:
                self.hwp.Run("TableRightCell")

        msg = f"✅ {filled}개 셀 채움"
        if skipped:
            msg += f"\n⚠️ 건너뜀: {', '.join(skipped[:5])}"
        return msg

    def format_table_cells(self, table_index: int, format_dict: dict) -> str:
        """
        현재 선택된 셀 범위의 테두리/배경을 변경합니다.
        먼저 대상 셀을 find_text_anchor 등으로 선택한 후 호출하세요.
        ActionTable: CellFill + CellBorderFill ParameterSet
          FillColor(RGB int), BorderTypeLeft/Right/Top/Bottom
        """
        self._ensure()
        try:
            if "fill_color" in format_dict:
                act = self.hwp.CreateAction("CellFill")
                ps = act.CreateSet()
                ps.SetItem("FillColor", int(format_dict["fill_color"]))
                act.Execute(ps)

            if "border_width" in format_dict:
                act = self.hwp.CreateAction("CellBorderFill")
                ps = act.CreateSet()
                for side in ("Left", "Right", "Top", "Bottom"):
                    key = f"border_type_{side.lower()}"
                    if key in format_dict:
                        ps.SetItem(f"BorderType{side}", format_dict[key])
                ps.SetItem("BorderWidth", format_dict["border_width"])
                act.Execute(ps)
            return "✅ 셀 서식 적용"
        except Exception as e:
            return f"❌ 셀 서식 실패: {e}"

    # ──────────────────────────────────────────────
    # 카테고리 4: Styling (서식 유지/미세 조정)
    # ──────────────────────────────────────────────

    def auto_fit_paragraph(self, decrease_count: int = 3) -> str:
        """
        자간을 좁혀 텍스트를 한 줄에 맞춥니다.
        ActionTable: CharShapeSpacingDecrease (Run 가능, ParameterSet 없음)
        """
        self._ensure()
        for _ in range(decrease_count):
            self.hwp.Run("CharShapeSpacingDecrease")
        return f"✅ 자간 {decrease_count}단계 축소"

    def set_font_style(self, font_name: str = None, size_pt: float = None,
                       bold: bool = None, color_rgb: int = None) -> str:
        """
        현재 선택 영역의 글자 서식을 설정합니다.
        ActionTable: CharShape ParameterSet
          Height(1/100pt 단위: 10pt=1000), FaceName, Bold, TextColor(RGB int)
        """
        self._ensure()
        try:
            act = self.hwp.CreateAction("CharShape")
            ps = act.CreateSet()
            if size_pt is not None:
                ps.SetItem("Height", int(size_pt * 100))
            if font_name:
                ps.SetItem("FaceName", font_name)
                for lang_key in ("FaceNameHangul", "FaceNameLatin", "FaceNameHanja"):
                    ps.SetItem(lang_key, font_name)
            if bold is not None:
                ps.SetItem("Bold", bold)
            if color_rgb is not None:
                ps.SetItem("TextColor", color_rgb)
            act.Execute(ps)
            return "✅ 글자 서식 적용"
        except Exception as e:
            return f"❌ 글자 서식 실패: {e}"

    # ──────────────────────────────────────────────
    # 카테고리 5: Structural Editing (문서 구조 편집)
    # ──────────────────────────────────────────────

    def append_page_from_template(self, file_path: str) -> str:
        """
        다른 HWP 파일을 현재 문서 끝에 이어 붙입니다.
        ActionTable: InsertFile ParameterSet
        """
        self._ensure()
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            return f"❌ 파일 없음: {abs_path}"
        try:
            self.hwp.Run("MoveDocEnd")
            act = self.hwp.CreateAction("InsertFile")
            ps = act.CreateSet()
            ps.SetItem("filename", abs_path)
            ps.SetItem("KeepSection", False)
            ps.SetItem("KeepCharshape", True)
            ps.SetItem("KeepParashape", True)
            act.Execute(ps)
            return f"✅ 파일 병합: {os.path.basename(abs_path)}"
        except Exception as e:
            return f"❌ 파일 병합 실패: {e}"

    def manage_page_visibility(self, page_number: int, action: str = "hide") -> str:
        """
        특정 페이지를 감추거나 보이게 합니다.
        ActionTable: PageHiding ParameterSet
        action: "hide" | "show"
        """
        self._ensure()
        try:
            act = self.hwp.CreateAction("PageHiding")
            ps = act.CreateSet()
            ps.SetItem("PageNum", page_number)
            ps.SetItem("Hide", action == "hide")
            act.Execute(ps)
            return f"✅ 페이지 {page_number} {'감춤' if action == 'hide' else '표시'}"
        except Exception as e:
            return f"❌ 페이지 처리 실패: {e}"

    # ──────────────────────────────────────────────
    # 카테고리 6: Validation & Export
    # ──────────────────────────────────────────────

    def check_missing_fields(self) -> dict:
        """
        채워지지 않은 필드와 플레이스홀더({{...}})를 검사합니다.
        ActionTable: ForwardFind + FindReplace ParameterSet
        """
        self._ensure()
        missing_fields = []
        for f in self.get_field_info():
            if not f["value"].strip():
                missing_fields.append(f["name"])

        # {{...}} 플레이스홀더 검색
        placeholders_found = False
        try:
            act = self.hwp.CreateAction("ForwardFind")
            ps = act.CreateSet()
            ps.SetItem("FindString", "{{")
            ps.SetItem("IgnoreCase", False)
            self.hwp.Run("MoveDocBegin")
            placeholders_found = bool(act.Execute(ps))
        except Exception:
            pass

        return {
            "missing_fields": missing_fields,
            "has_placeholders": placeholders_found,
            "ok": not missing_fields and not placeholders_found,
        }

    def export_to_pdf(self, target_path: str) -> str:
        """
        문서를 PDF로 저장합니다.
        ActionTable: PrintToPDF + Print ParameterSet (Range=3 전체)
        """
        self._ensure()
        abs_path = os.path.abspath(target_path)
        try:
            # 방법 1: PrintToPDF Action
            act = self.hwp.CreateAction("PrintToPDF")
            ps = act.CreateSet()
            ps.SetItem("filename", abs_path)
            ps.SetItem("Range", 3)
            ps.SetItem("Copies", 1)
            act.Execute(ps)
            return f"✅ PDF 저장: {abs_path}"
        except Exception:
            try:
                # 방법 2: SaveAs API
                self.hwp.SaveAs(abs_path, "PDF", "")
                return f"✅ PDF 저장: {abs_path}"
            except Exception as e:
                return f"❌ PDF 저장 실패: {e}"

    def save_as_hwp(self, target_path: str, distribute_mode: bool = False) -> str:
        """
        HWP로 저장합니다. distribute_mode=True면 배포용(수정 불가) 문서로 저장.
        ActionTable: FileSetSecurity* ParameterSet (CreateSet+Execute 필수)
        """
        self._ensure()
        if distribute_mode:
            try:
                act = self.hwp.CreateAction("FileSetSecurity")
                ps = act.CreateSet()
                ps.SetItem("NoPrint", False)
                ps.SetItem("NoCopy", False)
                ps.SetItem("NoModify", True)
                act.Execute(ps)
            except Exception as e:
                return f"❌ 배포용 설정 실패: {e}"
        return self.save(target_path)

    def finalize_document(self) -> str:
        """문서를 마무리합니다 (필드 잠금 해제, 저장)."""
        self._ensure()
        try:
            self.hwp.Run("FieldMerge")
        except Exception:
            pass
        return "✅ 문서 마무리 완료"

    # ──────────────────────────────────────────────
    # 카테고리 7: Raw Action Bridge
    # ──────────────────────────────────────────────

    def execute_raw_action(self, action_id: str, params: dict = None) -> str:
        """
        ActionTable.md의 모든 Action을 직접 실행합니다.
        params 없음(or '-' 표시): hwp.Run() 사용
        params 있음: CreateAction().CreateSet().Execute() 경로 사용
        '*' 표시 Action은 반드시 params와 함께 호출하세요.
        """
        self._ensure()
        try:
            if params:
                act = self.hwp.CreateAction(action_id)
                ps = act.CreateSet()
                for k, v in params.items():
                    ps.SetItem(k, v)
                result = bool(act.Execute(ps))
            else:
                result = bool(self.hwp.Run(action_id))
            return f"✅ {action_id}: {'성공' if result else '실패(False 반환)'}"
        except Exception as e:
            return f"❌ {action_id} 실패: {e}"

    # ──────────────────────────────────────────────
    # Tool Dispatcher (에이전틱 루프용)
    # ──────────────────────────────────────────────

    def dispatch_tool(self, tool_name: str, tool_args: dict) -> str:
        """
        LLM tool call을 해당 메서드로 라우팅합니다.
        main.py의 아ジェンtic 루프에서 메인 스레드로 호출됩니다.
        """
        try:
            if tool_name == "analyze_document_structure":
                return json.dumps(self.analyze_document_structure(), ensure_ascii=False)
            elif tool_name == "get_field_info":
                return json.dumps(self.get_field_info(), ensure_ascii=False)
            elif tool_name == "get_all_tables_overview":
                return json.dumps(self.get_all_tables_overview(), ensure_ascii=False)
            elif tool_name == "get_table_schema":
                return json.dumps(self.get_table_schema(tool_args["table_index"]), ensure_ascii=False)
            elif tool_name == "find_text_anchor":
                return json.dumps(self.find_text_anchor(tool_args["keyword"]), ensure_ascii=False)
            elif tool_name == "fill_field_data":
                return self.fill_field_data(tool_args["data_map"])
            elif tool_name == "replace_text_patterns":
                return self.replace_text_patterns(tool_args["mapping"])
            elif tool_name == "set_checkbox_state":
                return self.set_checkbox_state(tool_args["label"], tool_args["is_checked"])
            elif tool_name == "insert_image_box":
                return self.insert_image_box(
                    tool_args.get("anchor_text", ""),
                    tool_args["image_path"],
                    tool_args.get("size_mode", "fit"),
                )
            elif tool_name == "sync_table_rows":
                return self.sync_table_rows(
                    tool_args["table_index"], tool_args["data_count"],
                    tool_args.get("header_rows", 1),
                )
            elif tool_name == "fill_table_data_matrix":
                return self.fill_table_data_matrix(
                    tool_args["table_index"],
                    tool_args.get("start_row", 1),
                    tool_args["matrix"],
                    tool_args.get("cell_delay", 0.3),
                )
            elif tool_name == "format_table_cells":
                return self.format_table_cells(
                    tool_args["table_index"], tool_args.get("format_dict", {})
                )
            elif tool_name == "set_font_style":
                return self.set_font_style(
                    tool_args.get("font_name"),
                    tool_args.get("size_pt"),
                    tool_args.get("bold"),
                    tool_args.get("color_rgb"),
                )
            elif tool_name == "auto_fit_paragraph":
                return self.auto_fit_paragraph(tool_args.get("decrease_count", 3))
            elif tool_name == "append_page_from_template":
                return self.append_page_from_template(tool_args["file_path"])
            elif tool_name == "manage_page_visibility":
                return self.manage_page_visibility(
                    tool_args["page_number"], tool_args.get("action", "hide")
                )
            elif tool_name == "check_missing_fields":
                return json.dumps(self.check_missing_fields(), ensure_ascii=False)
            elif tool_name == "export_to_pdf":
                return self.export_to_pdf(tool_args["target_path"])
            elif tool_name == "save_as_hwp":
                return self.save_as_hwp(
                    tool_args["target_path"], tool_args.get("distribute_mode", False)
                )
            elif tool_name == "execute_raw_action":
                return self.execute_raw_action(
                    tool_args["action_id"], tool_args.get("params")
                )
            else:
                return f"❌ 알 수 없는 도구: {tool_name}"
        except KeyError as e:
            return f"❌ {tool_name} 필수 파라미터 누락: {e}"
        except Exception as e:
            return f"❌ {tool_name} 오류: {e}"

    # ──────────────────────────────────────────────
    # 레거시 호환 (기존 apply_edits 방식 유지)
    # ──────────────────────────────────────────────

    def apply_edits(self, edit_commands: list) -> list:
        """기존 JSON 명령 방식 호환. 신규 코드는 dispatch_tool을 사용하세요."""
        self._ensure()
        results = []
        for cmd in edit_commands:
            action = cmd.get("action", "")
            try:
                if action == "find_and_replace":
                    self._find_replace(cmd["find"], cmd["replace"],
                                       cmd.get("case_sensitive", False))
                    results.append(f"✅ find_and_replace: '{cmd['find'][:40]}' 교체")
                elif action == "replace_paragraph":
                    old = cmd.get("old_text") or cmd.get("find", "")
                    new = cmd.get("new_text") or cmd.get("replace", "")
                    if old:
                        self._find_replace(old, new, case_sensitive=True)
                        results.append(f"✅ replace_paragraph: '{old[:40]}' 교체")
                    else:
                        results.append("⚠️ replace_paragraph: old_text 없음")
                elif action == "set_field_value":
                    self.hwp.SetFieldText(cmd["field"], cmd["value"])
                    results.append(f"✅ set_field_value '{cmd['field']}'")
                else:
                    results.append(f"⚠️ 알 수 없는 action: '{action}'")
            except Exception as exc:
                results.append(f"❌ {action} 실패: {exc}")
        return results

    # ──────────────────────────────────────────────
    # 문서 읽기 (안정화 버전)
    # ──────────────────────────────────────────────

    def extract_document_structure(self) -> dict:
        """
        문서 구조 추출 (안정화 버전).
        1순위: GetTextFile("HTML") — 개선된 파서
        2순위: GetTextFile("UNICODE")
        3순위: InitScan(0x03) — 표 포함, 신뢰할 수 있는 옵션
        """
        self._ensure()
        log: list[str] = []
        paragraphs: list[dict] = []
        full_text = ""

        # ── 1순위: HTML ──────────────────────────────────────────────────
        html = ""
        try:
            html = self.hwp.GetTextFile("HTML", "") or ""
            log.append(f"GetTextFile('HTML'): {len(html)}자")
        except Exception as exc:
            log.append(f"GetTextFile('HTML') 실패: {exc}")

        if html and len(html) > 50:
            try:
                lines, paragraphs = self._parse_html_to_lines(html)
                full_text = "\n".join(lines)
                n_cells = sum(1 for l in lines if l.startswith("[셀"))
                log.append(f"HTML 파싱: {len(paragraphs)}개 단락, {n_cells}개 셀")
            except Exception as exc:
                log.append(f"HTML 파싱 실패: {exc}")
                full_text = ""
                paragraphs = []

        # ── 2순위: UNICODE ───────────────────────────────────────────────
        if not full_text:
            for fmt in ("UNICODE", "TEXT"):
                try:
                    candidate = self.hwp.GetTextFile(fmt, "") or ""
                    log.append(f"GetTextFile('{fmt}'): {len(candidate)}자")
                    if len(candidate) > len(full_text):
                        full_text = candidate
                    if full_text:
                        break
                except Exception as exc:
                    log.append(f"GetTextFile('{fmt}') 실패: {exc}")

        # ── 3순위: InitScan(0x03) — 표 포함, 안정적 옵션만 사용 ──────────
        if not full_text:
            try:
                self.hwp.InitScan(option=0x03)
                buf: list[str] = []
                pi = 0
                for _ in range(8000):
                    state, text = self.hwp.GetText()
                    tok = (text or "").strip()
                    if state == 0:
                        if buf:
                            t = " ".join(buf)
                            paragraphs.append({"index": pi, "text": t})
                        break
                    if state in (1, 2, 3):
                        if tok:
                            buf.append(tok)
                    if state in (2, 3):
                        t = " ".join(buf).strip()
                        if t:
                            paragraphs.append({"index": pi, "text": t})
                            pi += 1
                        buf = []
                if paragraphs:
                    full_text = "\n".join(p["text"] for p in paragraphs)
                    log.append(f"InitScan(0x03): {len(paragraphs)}개 단락")
            except Exception as exc:
                log.append(f"InitScan(0x03) 실패: {exc}")
            finally:
                try:
                    self.hwp.ReleaseScan()
                except Exception:
                    pass

        if not paragraphs and full_text:
            paragraphs = [
                {"index": i, "text": ln}
                for i, ln in enumerate(full_text.splitlines())
                if ln.strip()
            ]

        self._cached_paragraphs = paragraphs

        fields: list[str] = self._get_field_names()

        return {
            "total_paragraphs": len(paragraphs),
            "paragraphs": paragraphs,
            "fields": fields,
            "full_text": full_text,
            "extraction_log": log,
        }

    def get_summary_text(self, max_chars: int = 4000,
                         _structure: dict | None = None) -> str:
        s = _structure or self.extract_document_structure()
        header = f"[문서 요약]\n총 단락/셀 수: {s['total_paragraphs']}\n"
        if s["fields"]:
            header += f"누름틀: {', '.join(s['fields'])}\n"
        header += "\n[단락/셀 목록]\n"
        for p in s["paragraphs"]:
            if p["text"]:
                header += f"  [{p['index']}] {p['text']}\n"
        header += "\n[전체 텍스트]\n"
        combined = header + s["full_text"]
        if len(combined) <= max_chars:
            return combined
        kf = int(max_chars * 0.7)
        kb = max_chars - kf
        return combined[:kf] + f"\n\n...[{len(combined)-max_chars}자 생략]...\n\n" + combined[-kb:]

    # ──────────────────────────────────────────────
    # 내부 유틸리티
    # ──────────────────────────────────────────────

    def _find_replace(self, find: str, replace: str, case_sensitive: bool = False):
        """
        AllReplace 실행. ActionTable: AllReplace + FindReplace* ParameterSet.

        반드시 AllReplace 전에 텍스트 존재 여부를 확인합니다.
        존재하지 않는 텍스트로 AllReplace를 호출하면 HWP가
        "찾는 내용이 없습니다" 다이얼로그를 띄우고 COM 호출이 블로킹됩니다.
        """
        # 사전 확인: HTML에서 텍스트가 존재하지 않으면 HWP 호출 자체를 건너뜀
        if not self._text_exists(find, case_sensitive):
            return

        act = self.hwp.CreateAction("AllReplace")
        ps = act.CreateSet()
        ps.SetItem("FindString", find)
        ps.SetItem("ReplaceString", replace)
        ps.SetItem("IgnoreCase", not case_sensitive)
        ps.SetItem("WholeWordOnly", False)
        ps.SetItem("AutoSpell", True)
        ps.SetItem("Direction", 3)
        ps.SetItem("FindJaso", False)
        ps.SetItem("FindRegExp", False)
        ps.SetItem("ReplaceMode", 2)
        act.Execute(ps)
        self._cached_paragraphs = []

    def _text_exists(self, text: str, case_sensitive: bool = False) -> bool:
        """
        HTML 텍스트에서 검색어 존재 여부를 빠르게 확인합니다.
        AllReplace/ForwardFind 전에 호출해 HWP 다이얼로그 블로킹을 방지합니다.
        """
        html = self._get_html()
        if not html:
            return False
        # 공백 정규화 후 비교 (HTML의 \r\n\t vs 실제 텍스트 불일치 처리)
        normalized_html = " ".join(html.split())
        normalized_find = " ".join(text.split())
        if case_sensitive:
            return normalized_find in normalized_html
        return normalized_find.lower() in normalized_html.lower()

    def _get_field_names(self) -> list[str]:
        try:
            raw = self.hwp.GetFieldList(reserved=0) or ""
            return [n.strip() for n in raw.split("\x02") if n.strip()]
        except Exception:
            return []

    def _get_html(self) -> str:
        try:
            return self.hwp.GetTextFile("HTML", "") or ""
        except Exception:
            return ""

    def _parse_html_to_lines(self, html: str) -> tuple[list[str], list[dict]]:
        """HTML을 텍스트 라인과 단락 목록으로 파싱합니다 (colspan/rowspan 인식)."""

        class _Parser(HTMLParser):
            def __init__(self):
                super().__init__()
                self.lines: list[str] = []
                self.paragraphs: list[dict] = []
                self._pi = 0
                self._in_cell = False
                self._table_depth = 0
                self._cell_cnt = 0
                self._cbuf: list[str] = []
                self._tbuf: list[str] = []

            def handle_starttag(self, tag, attrs):
                t = tag.lower()
                if t == "table":
                    self._flush_text()
                    self._table_depth += 1
                    if self._table_depth == 1:
                        self._cell_cnt = 0
                        self.lines.append("[표 시작]")
                elif t in ("td", "th") and self._table_depth == 1:
                    self._in_cell = True
                    self._cbuf = []
                elif t == "br":
                    (self._cbuf if self._in_cell else self._tbuf).append(" ")
                elif t == "p":
                    self._flush_text()

            def handle_endtag(self, tag):
                t = tag.lower()
                if t == "table":
                    self._table_depth -= 1
                    if self._table_depth == 0:
                        self.lines.append("[표 끝]")
                elif t in ("td", "th") and self._table_depth == 1:
                    cell = "".join(self._cbuf).strip()
                    if cell:
                        self.lines.append(f"[셀{self._cell_cnt}] {cell}")
                        self.paragraphs.append({"index": self._pi, "text": cell})
                        self._pi += 1
                        self._cell_cnt += 1
                    self._in_cell = False
                    self._cbuf = []
                elif t == "p":
                    self._flush_text()

            def handle_data(self, data):
                if self._in_cell:
                    self._cbuf.append(data)
                elif data.strip():
                    self._tbuf.append(data.strip())

            def _flush_text(self):
                t = "".join(self._tbuf).strip()
                if t:
                    self.lines.append(t)
                    self.paragraphs.append({"index": self._pi, "text": t})
                    self._pi += 1
                self._tbuf = []

        p = _Parser()
        p.feed(html)
        p._flush_text()
        return p.lines, p.paragraphs

    def _parse_all_tables(self, html: str) -> list[list[list[str]]]:
        """HTML에서 모든 표의 셀 데이터를 추출합니다."""

        class _TableParser(HTMLParser):
            def __init__(self):
                super().__init__()
                self.tables: list[list[list[str]]] = []
                self._depth = 0
                self._cur_table: list[list[str]] | None = None
                self._cur_row: list[str] | None = None
                self._in_cell = False
                self._cbuf: list[str] = []

            def handle_starttag(self, tag, attrs):
                t = tag.lower()
                if t == "table":
                    self._depth += 1
                    if self._depth == 1:
                        self._cur_table = []
                elif t == "tr" and self._depth == 1:
                    self._cur_row = []
                elif t in ("td", "th") and self._depth == 1:
                    self._in_cell = True
                    self._cbuf = []
                elif t == "br" and self._in_cell:
                    self._cbuf.append(" ")

            def handle_endtag(self, tag):
                t = tag.lower()
                if t == "table":
                    if self._depth == 1 and self._cur_table is not None:
                        self.tables.append(self._cur_table)
                        self._cur_table = None
                    self._depth -= 1
                elif t == "tr" and self._depth == 1:
                    if self._cur_table is not None and self._cur_row is not None:
                        self._cur_table.append(self._cur_row)
                    self._cur_row = None
                elif t in ("td", "th") and self._depth == 1:
                    cell = "".join(self._cbuf).strip()
                    if self._cur_row is not None:
                        self._cur_row.append(cell)
                    self._in_cell = False
                    self._cbuf = []

            def handle_data(self, data):
                if self._in_cell:
                    self._cbuf.append(data)

        p = _TableParser()
        try:
            p.feed(html)
        except Exception:
            pass
        return p.tables

    def _ensure(self):
        if not self._connected or self.hwp is None:
            raise RuntimeError("한글에 연결되지 않았습니다. '한글에서 열기'를 먼저 실행하세요.")
