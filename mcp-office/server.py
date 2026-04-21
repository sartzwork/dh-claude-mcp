VERSION = "1.0.1"
"""
mcp-office / server.py
Excel / Word / PowerPoint 문서를 COM으로 직접 읽는 MCP 서버
import mcp 없이 stdio JSON으로 직접 MCP 프로토콜 구현
"""

import sys
import json

# ════════════════════════════════════════════════════════════════════
# 자동 업데이트
# GitHub raw URL 에서 버전을 확인하고 최신 버전이면 server.py 교체 후 재시작
# 업데이트 실패 시 기존 버전으로 계속 실행
# ════════════════════════════════════════════════════════════════════

_SERVER_URL = "https://raw.githubusercontent.com/sartzwork/dh-claude-mcp/main/mcp-office/server.py"

def _check_and_update():
    try:
        import urllib.request, os, subprocess
        req = urllib.request.Request(_SERVER_URL, headers={"Range": "bytes=0-30"})
        with urllib.request.urlopen(req, timeout=3) as r:
            head = r.read().decode("utf-8", errors="replace")
        latest = ""
        for line in head.splitlines():
            if line.startswith("VERSION"):
                latest = line.split('"')[1]
                break
        if latest and latest != VERSION:
            with urllib.request.urlopen(_SERVER_URL, timeout=10) as r:
                new_code = r.read()
            with open(__file__, "wb") as f:
                f.write(new_code)
            # Windows에서는 os.execv 미지원 → subprocess로 재시작
            subprocess.Popen([sys.executable] + sys.argv)
            sys.exit(0)
    except Exception:
        pass  # 업데이트 실패 시 기존 버전으로 계속 실행

_check_and_update()

# ══════════════════════════════════════════════════════════════
# 공통 유틸
# ══════════════════════════════════════════════════════════════

def make_ok(data: dict) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2)

def make_err(msg: str) -> str:
    return make_ok({"success": False, "error": msg})


# ══════════════════════════════════════════════════════════════
# Excel 헬퍼
# ══════════════════════════════════════════════════════════════

def get_excel_app():
    import win32com.client
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        raise RuntimeError("Microsoft Excel이 실행 중이지 않습니다. Excel을 먼저 열어주세요.")

def get_active_workbook(excel):
    try:
        wb = excel.ActiveWorkbook
        if wb is None:
            raise RuntimeError("열린 통합 문서가 없습니다.")
        return wb
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"활성 통합 문서를 가져올 수 없습니다: {e}")

def col_index_to_letter(col_idx: int) -> str:
    result = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(ord("A") + rem) + result
    return result

def extract_sheet_texts(wb, max_rows: int = 500, target_sheet_name: str = None):
    sheets_data = []
    sheets_iter = [wb.Sheets(target_sheet_name)] if target_sheet_name else list(wb.Sheets)
    for sheet in sheets_iter:
        try:
            used_range = sheet.UsedRange
            if used_range is None:
                continue
            row_count = used_range.Rows.Count
            col_count = used_range.Columns.Count
            capped_rows = min(row_count, max_rows)
            start_row = used_range.Row
            start_col = used_range.Column
            if capped_rows < row_count:
                capped_range = sheet.Range(
                    sheet.Cells(start_row, start_col),
                    sheet.Cells(start_row + capped_rows - 1, start_col + col_count - 1)
                )
                values = capped_range.Value
            else:
                values = used_range.Value
            if row_count == 1 and col_count == 1:
                values = ((values,),)
            elif row_count == 1:
                values = (values,)
            elif col_count == 1:
                values = tuple((v,) for v in values)
            cells_data = []
            for r_offset, row_vals in enumerate(values):
                if row_vals is None:
                    continue
                for c_offset, val in enumerate(row_vals):
                    # 문자열뿐 아니라 숫자, 날짜 등 모든 값 포함 (데이터 불일치 확인용)
                    if val is not None and str(val).strip():
                        col_letter = col_index_to_letter(start_col + c_offset)
                        row_num = start_row + r_offset
                        cells_data.append({"cell": f"{col_letter}{row_num}", "text": str(val).strip()})
            if cells_data:
                sheets_data.append({"sheet_name": sheet.Name, "cells": cells_data})
        except Exception:
            continue
    return sheets_data


# ══════════════════════════════════════════════════════════════
# Word 헬퍼
# ══════════════════════════════════════════════════════════════

def get_word_app():
    import win32com.client
    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception:
        raise RuntimeError("Microsoft Word가 실행 중이지 않습니다. Word를 먼저 열어주세요.")

def get_active_doc(word):
    try:
        doc = word.ActiveDocument
        if doc is None:
            raise RuntimeError("열린 문서가 없습니다.")
        return doc
    except Exception as e:
        raise RuntimeError(f"활성 문서를 가져올 수 없습니다: {e}")


# ══════════════════════════════════════════════════════════════
# PowerPoint 헬퍼
# ══════════════════════════════════════════════════════════════

def get_ppt_app():
    import win32com.client
    try:
        return win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        raise RuntimeError("Microsoft PowerPoint가 실행 중이지 않습니다. PowerPoint를 먼저 열어주세요.")

def get_active_presentation(ppt):
    try:
        pres = ppt.ActivePresentation
        if pres is None:
            raise RuntimeError("열린 프레젠테이션이 없습니다.")
        return pres
    except Exception as e:
        raise RuntimeError(f"활성 프레젠테이션을 가져올 수 없습니다: {e}")

def extract_slide_texts(pres):
    slides_data = []
    for slide_idx, slide in enumerate(pres.Slides):
        slide_texts = []
        for shape in slide.Shapes:
            try:
                if shape.HasTextFrame:
                    for para in shape.TextFrame.TextRange.Paragraphs():
                        text = para.Text.strip()
                        if text:
                            slide_texts.append({"shape_name": shape.Name, "text": text})
            except Exception:
                continue
        if slide_texts:
            slides_data.append({
                "slide_number": slide_idx + 1,
                "slide_name": slide.Name,
                "texts": slide_texts
            })
    return slides_data


# ══════════════════════════════════════════════════════════════
# 툴 목록
# ══════════════════════════════════════════════════════════════

TOOLS = [
    # Excel
    {"name": "excel_get_status", "description": "현재 Excel 연결 상태와 열린 통합 문서 목록을 반환합니다.", "inputSchema": {"type": "object", "properties": {}, "required": []}},
    {"name": "excel_get_text", "description": "현재 열린 Excel 파일의 모든 시트에서 텍스트가 있는 셀을 가져옵니다.", "inputSchema": {"type": "object", "properties": {"sheet_name": {"type": "string", "description": "특정 시트만 읽을 경우 시트 이름 (생략 시 전체)"}, "max_rows": {"type": "integer", "description": "시트당 최대 읽을 행 수 (기본값: 500)", "default": 500}}, "required": []}},
    {"name": "excel_proofread", "description": "현재 열린 Excel 파일의 텍스트 셀에서 오탈자·어색한 표현을 AI로 교정합니다.", "inputSchema": {"type": "object", "properties": {"focus": {"type": "string", "enum": ["오탈자", "맞춤법", "문법", "어색한표현"], "description": "집중 교정 유형 (생략 시 전체)"}, "sheet_name": {"type": "string", "description": "특정 시트만 교정 (생략 시 전체)"}, "max_rows": {"type": "integer", "default": 500}}, "required": []}},
    # Word
    {"name": "word_get_status", "description": "현재 Word 연결 상태와 열린 문서 목록을 반환합니다.", "inputSchema": {"type": "object", "properties": {}, "required": []}},
    {"name": "word_get_text", "description": "현재 열린 Word 문서의 전체 텍스트를 문단 단위로 가져옵니다.", "inputSchema": {"type": "object", "properties": {}, "required": []}},
    {"name": "word_proofread", "description": "현재 열린 Word 문서의 오탈자·어색한 표현을 AI로 교정합니다.", "inputSchema": {"type": "object", "properties": {"focus": {"type": "string", "enum": ["오탈자", "맞춤법", "문법", "어색한표현"], "description": "집중 교정 유형 (생략 시 전체)"}}, "required": []}},
    # PowerPoint
    {"name": "ppt_get_status", "description": "현재 PowerPoint 연결 상태와 열린 프레젠테이션 목록을 반환합니다.", "inputSchema": {"type": "object", "properties": {}, "required": []}},
    {"name": "ppt_get_text", "description": "현재 열린 PowerPoint 파일의 모든 슬라이드 텍스트를 슬라이드 번호와 함께 가져옵니다.", "inputSchema": {"type": "object", "properties": {}, "required": []}},
    {"name": "ppt_proofread", "description": "현재 열린 PowerPoint 파일의 오탈자·어색한 표현을 AI로 교정합니다.", "inputSchema": {"type": "object", "properties": {"focus": {"type": "string", "enum": ["오탈자", "맞춤법", "문법", "어색한표현"], "description": "집중 교정 유형 (생략 시 전체)"}}, "required": []}},
]


# ══════════════════════════════════════════════════════════════
# 툴 실행
# ══════════════════════════════════════════════════════════════

def handle_tool(name: str, arguments: dict) -> str:

    # Excel
    if name == "excel_get_status":
        try:
            excel = get_excel_app()
            workbooks = [{"name": wb.Name, "path": wb.FullName, "sheet_count": wb.Sheets.Count, "is_active": (wb.Name == excel.ActiveWorkbook.Name)} for wb in excel.Workbooks]
            return make_ok({"success": True, "status": "connected", "open_workbooks": workbooks})
        except Exception as e:
            return make_ok({"success": False, "status": "disconnected", "error": str(e)})

    if name == "excel_get_text":
        try:
            excel = get_excel_app()
            wb = get_active_workbook(excel)
            sheets_data = extract_sheet_texts(wb, arguments.get("max_rows", 200), arguments.get("sheet_name"))
            return make_ok({"success": True, "workbook_name": wb.Name, "sheets": sheets_data})
        except Exception as e:
            return make_err(str(e))

    if name == "excel_proofread":
        try:
            excel = get_excel_app()
            wb = get_active_workbook(excel)
            sheets_data = extract_sheet_texts(wb, arguments.get("max_rows", 200), arguments.get("sheet_name"))
        except Exception as e:
            return make_err(str(e))
        if not sheets_data:
            return make_err("통합 문서에 교정할 텍스트가 없습니다.")
        focus = arguments.get("focus", "")
        focus_note = f" 특히 '{focus}' 유형에 집중해서 교정하세요." if focus else ""
        return make_ok({
            "success": True, "workbook_name": wb.Name, "focus": focus or "전체",
            "note": f"아래 sheets 데이터를 바탕으로 오탈자·맞춤법 오류·어색한 표현·수치 및 데이터 불일치를 교정해주세요.{focus_note}",
            "sheets": sheets_data,
            "notice": "교정 결과는 참고용입니다. 문서 수정은 사용자가 직접 Excel에서 진행해 주세요."
        })

    # Word
    if name == "word_get_status":
        try:
            word = get_word_app()
            docs = [{"name": doc.Name, "path": doc.FullName, "is_active": (doc.Name == word.ActiveDocument.Name)} for doc in word.Documents]
            return make_ok({"success": True, "status": "connected", "open_documents": docs})
        except Exception as e:
            return make_ok({"success": False, "status": "disconnected", "error": str(e)})

    if name == "word_get_text":
        try:
            word = get_word_app()
            doc = get_active_doc(word)
            full_text = doc.Content.Text
            paragraphs = [{"index": i, "text": t.rstrip("\r")} for i, t in enumerate(full_text.split("\r")) if t.strip()]
            return make_ok({"success": True, "document_name": doc.Name, "paragraph_count": len(paragraphs), "paragraphs": paragraphs})
        except Exception as e:
            return make_err(str(e))

    if name == "word_proofread":
        try:
            word = get_word_app()
            doc = get_active_doc(word)
        except Exception as e:
            return make_err(str(e))
        full_text = doc.Content.Text
        paragraphs = [{"index": i, "text": t.rstrip("\r")} for i, t in enumerate(full_text.split("\r")) if t.strip()]
        if not paragraphs:
            return make_err("문서에 교정할 텍스트가 없습니다.")
        full_text = "\n".join([f"[{p['index']}] {p['text']}" for p in paragraphs])
        focus = arguments.get("focus", "")
        focus_note = f"\n특히 '{focus}' 유형에 집중해서 교정하세요." if focus else ""
        return (
            f"다음은 Microsoft Word 문서의 내용입니다. 각 줄은 [문단번호] 텍스트 형식입니다.\n"
            f"오탈자, 맞춤법 오류, 어색한 표현을 찾아 교정 목록을 JSON으로 반환하세요.{focus_note}\n\n"
            f"반드시 아래 형식만 출력하세요 (설명·마크다운 없이 순수 JSON):\n"
            f'{{"corrections":[{{"original":"원본 텍스트","corrected":"수정된 텍스트","reason":"수정 이유"}}]}}\n\n'
            f"문서 내용:\n{full_text}\n\n"
            f"---\n교정 결과는 참고용입니다. 문서 수정은 사용자가 직접 Word에서 진행해 주세요."
        )

    # PowerPoint
    if name == "ppt_get_status":
        try:
            ppt = get_ppt_app()
            presentations = [{"name": pres.Name, "path": pres.FullName, "slide_count": pres.Slides.Count, "is_active": (pres.Name == ppt.ActivePresentation.Name)} for pres in ppt.Presentations]
            return make_ok({"success": True, "status": "connected", "open_presentations": presentations})
        except Exception as e:
            return make_ok({"success": False, "status": "disconnected", "error": str(e)})

    if name == "ppt_get_text":
        try:
            ppt = get_ppt_app()
            pres = get_active_presentation(ppt)
            slides_data = extract_slide_texts(pres)
            return make_ok({"success": True, "presentation_name": pres.Name, "slide_count": pres.Slides.Count, "slides": slides_data})
        except Exception as e:
            return make_err(str(e))

    if name == "ppt_proofread":
        try:
            ppt = get_ppt_app()
            pres = get_active_presentation(ppt)
        except Exception as e:
            return make_err(str(e))
        slides_data = extract_slide_texts(pres)
        if not slides_data:
            return make_err("프레젠테이션에 교정할 텍스트가 없습니다.")
        lines = [f"[슬라이드{s['slide_number']}][{item['shape_name']}] {item['text']}" for s in slides_data for item in s["texts"]]
        full_text = "\n".join(lines)
        focus = arguments.get("focus", "")
        focus_note = f"\n특히 '{focus}' 유형에 집중해서 교정하세요." if focus else ""
        return (
            f"다음은 Microsoft PowerPoint 프레젠테이션의 텍스트입니다.\n"
            f"각 줄은 [슬라이드번호][도형이름] 텍스트 형식입니다.\n"
            f"오탈자, 맞춤법 오류, 어색한 표현을 찾아 교정 목록을 JSON으로 반환하세요.{focus_note}\n\n"
            f"반드시 아래 형식만 출력하세요 (설명·마크다운 없이 순수 JSON):\n"
            f'{{"corrections":[{{"slide":1,"shape":"도형이름","original":"원본 텍스트","corrected":"수정된 텍스트","reason":"수정 이유"}}]}}\n\n'
            f"프레젠테이션 내용:\n{full_text}\n\n"
            f"---\n교정 결과는 참고용입니다. 문서 수정은 사용자가 직접 PowerPoint에서 진행해 주세요."
        )

    return make_err(f"알 수 없는 툴: {name}")


# ══════════════════════════════════════════════════════════════
# MCP 프로토콜 - stdio JSON 직접 구현
# ══════════════════════════════════════════════════════════════

def send(obj: dict):
    """JSON-RPC 응답을 stdout 으로 전송 (UTF-8 강제)"""
    data = json.dumps(obj, ensure_ascii=False) + "\n"
    sys.stdout.buffer.write(data.encode("utf-8"))
    sys.stdout.buffer.flush()


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except json.JSONDecodeError:
            continue

        req_id = req.get("id")
        method = req.get("method", "")

        # initialize
        if method == "initialize":
            send({
                "jsonrpc": "2.0", "id": req_id,
                "result": {
                    "protocolVersion": "2025-11-25",
                    "capabilities": {"tools": {}},
                    "serverInfo": {"name": "mcp-office", "version": "1.0.0"},
                }
            })

        # notifications/initialized (응답 불필요)
        elif method == "notifications/initialized":
            pass

        # tools/list
        elif method == "tools/list":
            send({
                "jsonrpc": "2.0", "id": req_id,
                "result": {"tools": TOOLS}
            })

        # tools/call
        elif method == "tools/call":
            params = req.get("params", {})
            tool_name = params.get("name", "")
            arguments = params.get("arguments", {})
            result_text = handle_tool(tool_name, arguments)
            send({
                "jsonrpc": "2.0", "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": result_text}]
                }
            })

        # 기타 알 수 없는 메서드
        else:
            if req_id is not None:
                send({
                    "jsonrpc": "2.0", "id": req_id,
                    "error": {"code": -32601, "message": f"Method not found: {method}"}
                })


if __name__ == "__main__":
    main()
