"""
mcp-hwp / server.py
열려있는 한글(HWP) 문서의 오탈자·어색한 표현·데이터 불일치를 찾아주는 MCP 서버
import mcp 없이 stdio JSON으로 직접 MCP 프로토콜 구현
"""

import sys
import json
import time

# ════════════════════════════════════════════════════════════════════
# 자동 업데이트
# GitHub raw URL 에서 버전을 확인하고 최신 버전이면 server.py 교체 후 재시작
# 업데이트 실패 시 기존 버전으로 계속 실행
# ════════════════════════════════════════════════════════════════════

VERSION = "1.0.0"
_BASE_URL = "https://raw.githubusercontent.com/sartzwork/dh-claude-mcp/main/mcp-hwp"

def _check_and_update():
    try:
        import urllib.request, os
        with urllib.request.urlopen(_BASE_URL + "/version.txt", timeout=3) as r:
            latest = r.read().decode().strip()
        if latest != VERSION:
            with urllib.request.urlopen(_BASE_URL + "/server.py", timeout=10) as r:
                new_code = r.read()
            with open(__file__, "wb") as f:
                f.write(new_code)
            os.execv(sys.executable, [sys.executable] + sys.argv)
    except Exception:
        pass  # 업데이트 실패 시 기존 버전으로 계속 실행

_check_and_update()

# ── pywin32 가용성 확인 ──────────────────────────────────────────────
try:
    import win32com.client as win32
    import pythoncom
    import win32clipboard
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


# ════════════════════════════════════════════════════════════════════
# HWP 연결 및 텍스트 추출
# ════════════════════════════════════════════════════════════════════

def _connect_hwp():
    """ROT 메모리에서 열려있는 한글 객체에 연결"""
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()
    for moniker in rot:
        name = moniker.GetDisplayName(context, None)
        if "hwp" in name.lower() or "hanword" in name.lower():
            obj = rot.GetObject(moniker)
            return win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    return None


def _extract_text(hwp) -> str:
    """클립보드 우회 방식으로 전체 텍스트 추출 (기존 클립보드 보존)"""

    # 기존 클립보드 백업
    original_clipboard = ""
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            original_clipboard = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
    except Exception:
        try: win32clipboard.CloseClipboard()
        except: pass

    extracted = ""
    try:
        # 지연 로딩 방지: 문서 끝→처음 이동으로 전체 렌더링 강제
        hwp.HAction.Run("MoveDocEnd")
        time.sleep(0.3)
        hwp.HAction.Run("MoveDocBegin")
        time.sleep(0.3)

        # 클립보드 초기화
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
        except Exception:
            pass

        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Copy")

        # 최대 5초 폴링
        for _ in range(25):
            time.sleep(0.2)
            try:
                win32clipboard.OpenClipboard()
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                    text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    if text:
                        extracted = text
                        win32clipboard.CloseClipboard()
                        break
                win32clipboard.CloseClipboard()
            except Exception:
                pass

        hwp.HAction.Run("Cancel")
        hwp.HAction.Run("MoveDocBegin")

    except Exception as e:
        try: win32clipboard.CloseClipboard()
        except: pass
        raise RuntimeError(f"텍스트 추출 실패: {e}")

    # 기존 클립보드 복원
    if original_clipboard:
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, original_clipboard)
            win32clipboard.CloseClipboard()
        except Exception:
            try: win32clipboard.CloseClipboard()
            except: pass

    return extracted.strip()


# ════════════════════════════════════════════════════════════════════
# 도구 목록
# ════════════════════════════════════════════════════════════════════

TOOLS = [
    {
        "name": "hwp_get_status",
        "description": "현재 실행 중인 한글(HWP) 프로세스 연결 상태를 반환합니다.",
        "inputSchema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "hwp_get_document",
        "description": "현재 한글에서 열려있는 문서의 파일명, 경로, 전체 텍스트(문단 목록)를 반환합니다.",
        "inputSchema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "hwp_proofread",
        "description": (
            "현재 열린 한글 문서의 텍스트를 추출한 뒤 AI가 오탈자·어색한 표현·"
            "데이터 불일치를 검토하여 교정 리포트를 반환합니다."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "focus": {
                    "type": "string",
                    "description": "검토 집중 항목. 예: 'typo'(오탈자), 'style'(어색한 표현), 'data'(수치·날짜 불일치), 'all'(전체, 기본값)",
                    "enum": ["typo", "style", "data", "all"],
                    "default": "all",
                }
            },
            "required": [],
        },
    },
]


# ════════════════════════════════════════════════════════════════════
# 도구 실행
# ════════════════════════════════════════════════════════════════════

def handle_tool(name: str, arguments: dict) -> str:

    if not WIN32_AVAILABLE:
        return "❌ pywin32 패키지가 설치되지 않았습니다.\n터미널에서 'pip install pywin32' 를 실행해 주세요."

    if name == "hwp_get_status":
        pythoncom.CoInitialize()
        try:
            hwp = _connect_hwp()
            if hwp:
                try:
                    doc_name = hwp.CurDocumentPath if hasattr(hwp, "CurDocumentPath") else "(알 수 없음)"
                except Exception:
                    doc_name = "(경로 확인 불가)"
                return f"✅ 한글 연결 성공\n현재 문서: {doc_name}"
            else:
                return "⚠️ 실행 중인 한글 프로세스를 찾을 수 없습니다."
        finally:
            pythoncom.CoUninitialize()

    elif name == "hwp_get_document":
        pythoncom.CoInitialize()
        try:
            hwp = _connect_hwp()
            if not hwp:
                return "⚠️ 한글 프로세스에 연결할 수 없습니다."
            text = _extract_text(hwp)
            if not text:
                return "⚠️ 문서 텍스트를 추출하지 못했습니다."
            paragraphs = [p for p in text.split("\n") if p.strip()]
            try:
                doc_path = hwp.CurDocumentPath
            except Exception:
                doc_path = "(경로 확인 불가)"
            return "\n".join([
                f"📄 파일 경로: {doc_path}",
                f"📝 총 문단 수: {len(paragraphs)}",
                f"📏 총 글자 수: {len(text):,}",
                "",
                "── 전체 텍스트 ──",
                text,
            ])
        finally:
            pythoncom.CoUninitialize()

    elif name == "hwp_proofread":
        focus = arguments.get("focus", "all")
        pythoncom.CoInitialize()
        try:
            hwp = _connect_hwp()
            if not hwp:
                return "⚠️ 한글 프로세스에 연결할 수 없습니다."
            text = _extract_text(hwp)
            if not text:
                return "⚠️ 문서 텍스트를 추출하지 못했습니다."
            focus_map = {
                "typo":  "오탈자(맞춤법·띄어쓰기·동음이의어 오용 등)만 집중 검토",
                "style": "어색한 표현(문장 구조, 중복 표현, 번역 투, 지나친 수동태 등)만 집중 검토",
                "data":  "수치·날짜·단위·고유명사 등 데이터 불일치나 논리 모순만 집중 검토",
                "all":   "오탈자, 어색한 표현, 데이터 불일치 전체 종합 검토",
            }
            focus_instruction = focus_map.get(focus, focus_map["all"])
            return (
                f"아래는 한글 문서에서 추출한 텍스트입니다.\n"
                f"검토 범위: {focus_instruction}\n\n"
                f"다음 형식으로 문제점을 정리해 주세요.\n"
                f"1. 오탈자 목록 (위치·원문·수정안)\n"
                f"2. 어색한 표현 목록 (원문·수정안·이유)\n"
                f"3. 데이터 불일치·논리 모순 목록 (위치·내용·이유)\n"
                f"4. 종합 의견\n\n"
                f"── 문서 원문 ──\n{text}\n\n"
                f"── 안내 ──\n"
                f"교정 결과는 참고용입니다. 문서 수정은 사용자가 직접 한글에서 진행해 주세요."
            )
        finally:
            pythoncom.CoUninitialize()

    else:
        return f"❌ 알 수 없는 도구: {name}"


# ════════════════════════════════════════════════════════════════════
# MCP 프로토콜 - stdio JSON 직접 구현
# ════════════════════════════════════════════════════════════════════

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
                    "serverInfo": {"name": "mcp-hwp", "version": "1.0.0"},
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
