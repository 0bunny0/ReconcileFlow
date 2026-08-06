from __future__ import annotations

import base64
from binascii import Error as Base64Error
import hashlib
import json
import math
from pathlib import Path
import re
from typing import Any

import streamlit as st

from reconciliation import (
    ReconciliationError,
    basis_config,
    export_excel,
    json_value,
    module_config,
    payload,
    reconcile_module_all,
)


APP_NAME = "POOWARD ReconcileFlow"
ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
MAX_COMBINED_UPLOAD_BYTES = 150 * 1024 * 1024

st.set_page_config(
    page_title=APP_NAME,
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def _component_markup() -> str:
    """Use the supplied index.html body without changing its visual structure."""
    source = (ASSETS / "index.html").read_text(encoding="utf-8")
    match = re.search(r"<body>(.*)</body>", source, flags=re.IGNORECASE | re.DOTALL)
    if not match:
        raise RuntimeError("assets/index.html 缺少 body 结构。")
    body = match.group(1)
    return re.sub(
        r"\s*<script\b[^>]*\bsrc=[\"']/app\.js[\"'][^>]*>\s*</script>\s*",
        "\n",
        body,
        flags=re.IGNORECASE,
    )


# This only removes Streamlit's outer shell. It does not alter the supplied UI.
STREAMLIT_SHELL_ADAPTER = r"""
#MainMenu,
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
footer[data-testid="stFooter"] {
  display: none !important;
}

[data-testid="stAppViewContainer"],
[data-testid="stMain"],
.stApp {
  background: var(--bg) !important;
}

[data-testid="stAppViewContainer"] {
  overflow-x: clip;
}

[data-testid="stMainBlockContainer"] {
  width: 100% !important;
  max-width: none !important;
  padding: 0 !important;
}

[data-testid="stMainBlockContainer"] > div,
[data-testid="stMainBlockContainer"] [data-testid="stVerticalBlock"],
[data-testid="stMainBlockContainer"] [data-testid="stElementContainer"] {
  width: 100% !important;
  max-width: none !important;
  gap: 0 !important;
  margin: 0 !important;
  padding: 0 !important;
}
"""


RECONCILE_UI = st.components.v2.component(
    "pooward_reconcile_flow_exact_ui",
    html=_component_markup(),
    css=(ASSETS / "styles.css").read_text(encoding="utf-8") + STREAMLIT_SHELL_ADAPTER,
    js=(ASSETS / "component.js").read_text(encoding="utf-8"),
    isolate_styles=False,
)


def _decode_file(item: Any, label: str) -> tuple[str, bytes]:
    if not isinstance(item, dict):
        raise ReconciliationError(f"请重新选择{label}。")
    name = str(item.get("name") or label).strip()
    if Path(name).suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ReconciliationError(f"{label}必须是 .xlsx 或 .xlsm 文件。")
    encoded = item.get("data")
    if not isinstance(encoded, str) or not encoded:
        raise ReconciliationError(f"没有收到{label}的文件内容。")
    try:
        content = base64.b64decode(encoded, validate=True)
    except (Base64Error, ValueError) as exc:
        raise ReconciliationError(f"{label}上传内容不完整，请重新选择。") from exc
    if not content:
        raise ReconciliationError(f"{label}是空文件。")
    return name, content


def _json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    converted = json_value(value)
    if isinstance(converted, float) and not math.isfinite(converted):
        return None
    return converted


def _success_response(request: dict[str, Any]) -> dict[str, Any]:
    request_id = str(request.get("request_id") or "")
    module = str(request.get("module") or "")
    module_info = module_config(module)
    try:
        tolerance = float(request.get("tolerance", 1.0))
    except (TypeError, ValueError) as exc:
        raise ReconciliationError("差异容差必须是数字。") from exc
    if tolerance < 0:
        raise ReconciliationError("差异容差不能小于 0。")

    files = request.get("files")
    if not isinstance(files, dict):
        raise ReconciliationError("请完整上传当前模块的三份 Excel 文件。")
    document_name, document = _decode_file(files.get("document"), "文控登记表")
    amount_name, amount = _decode_file(files.get("amount"), module_info["amount_source_label"])
    freight_name, freight = _decode_file(files.get("freight"), module_info["freight_source_label"])
    if len(document) + len(amount) + len(freight) > MAX_COMBINED_UPLOAD_BYTES:
        raise ReconciliationError("三份文件合计不能超过 150 MB。")

    source_names = {
        "document": document_name,
        "amount": amount_name,
        "freight": freight_name,
    }
    results = reconcile_module_all(
        document,
        amount,
        freight,
        module=module,
        tolerance=tolerance,
        source_names=source_names,
    )
    module_payload: dict[str, Any] = {}
    for basis, result in results.items():
        result_payload = payload(result)
        result_payload["download_name"] = (
            f"{module_info['label']}_{basis_config(result.basis)['label']}_核对结果.xlsx"
        )
        result_payload["download_base64"] = base64.b64encode(export_excel(result)).decode("ascii")
        module_payload[basis] = _json_safe(result_payload)
    return {
        "status": "ok",
        "request_id": request_id,
        "module": module,
        "document_digest": hashlib.sha256(document).hexdigest(),
        "modules": {module: module_payload},
    }


def _process_request(request: Any) -> dict[str, Any]:
    request_id = ""
    if isinstance(request, dict):
        request_id = str(request.get("request_id") or "")
    try:
        if not isinstance(request, dict) or not request_id:
            raise ReconciliationError("上传请求无效，请重新点击开始核对。")
        return _success_response(request)
    except ReconciliationError as exc:
        return {"status": "error", "request_id": request_id, "error": str(exc)}
    except Exception:
        return {
            "status": "error",
            "request_id": request_id,
            "error": "处理失败。请确认当前模块的三份 Excel 格式和表头未改变，然后重试。",
        }


def main() -> None:
    if "component_response" not in st.session_state:
        st.session_state.component_response = None
    if "processed_request_id" not in st.session_state:
        st.session_state.processed_request_id = None
    if "document_digest" not in st.session_state:
        st.session_state.document_digest = None
    if "module_payload_cache" not in st.session_state:
        st.session_state.module_payload_cache = {}

    result = RECONCILE_UI(
        data={"response": st.session_state.component_response},
        key="pooward_reconcile_flow",
        width="stretch",
        height="content",
        on_analyze_change=lambda: None,
    )
    request = getattr(result, "analyze", None)
    request_id = str(request.get("request_id") or "") if isinstance(request, dict) else ""
    if request_id and request_id != st.session_state.processed_request_id:
        st.session_state.processed_request_id = request_id
        response = _process_request(request)
        if response.get("status") == "ok":
            document_digest = response.get("document_digest")
            if (
                st.session_state.document_digest
                and st.session_state.document_digest != document_digest
            ):
                st.session_state.module_payload_cache = {}
            st.session_state.document_digest = document_digest
            cache = dict(st.session_state.module_payload_cache)
            cache.update(response.get("modules") or {})
            st.session_state.module_payload_cache = cache
            response["modules"] = cache
        st.session_state.component_response = response
        st.rerun()


if __name__ == "__main__":
    main()
