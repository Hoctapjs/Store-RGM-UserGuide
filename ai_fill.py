# -*- coding: utf-8 -*-
"""
Đọc ảnh chụp màn hình bằng Claude vision -> sinh content.yaml.

    python ai_fill.py shots/ ket-qua.yaml --topic "Quy trình 0502 bảo trì, vai trò cửa hàng"

Đặt tên file ảnh có số thứ tự để giữ đúng trình tự thao tác:
    01-tao-phieu.png, 02-chon-thiet-bi.png, ...

Cần biến môi trường ANTHROPIC_API_KEY, hoặc đã chạy `ant auth login`.
"""
import os, sys, io, json, base64, argparse, mimetypes

for _s in ("stdout", "stderr"):
    _st = getattr(sys, _s, None)
    if _st and hasattr(_st, "buffer"):
        setattr(sys, _s, io.TextIOWrapper(_st.buffer, encoding="utf-8",
                                          errors="replace", line_buffering=True))

import anthropic

MODEL = "claude-opus-5"
EXTS = (".png", ".jpg", ".jpeg", ".webp", ".gif")

# --------------------------------------------------------------------------
# Schema ràng buộc đầu ra - khớp đúng các layout render.py hiểu được.
# Nhờ structured output, model không thể trả về layout lạ hay thiếu trường.
# --------------------------------------------------------------------------
SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["slides"],
    "properties": {
        "slides": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["layout", "title"],
                "properties": {
                    "layout": {
                        "type": "string",
                        "enum": ["title", "section", "agenda", "two_col",
                                 "screenshot_callout", "screenshot_bullets",
                                 "process", "table"],
                    },
                    "title": {"type": "string"},
                    "subtitle": {"type": "string"},
                    "number": {"type": "string"},
                    "image": {"type": "string"},
                    "image_side": {"type": "string", "enum": ["left", "right"]},
                    "note": {"type": "string"},
                    "bullets": {"type": "array", "items": {"type": "string"}},
                    "items": {"type": "array", "items": {"type": "string"}},
                    "callouts": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "required": ["text"],
                            "properties": {
                                "text": {"type": "string"},
                                "note": {"type": "string"},
                            },
                        },
                    },
                    "columns": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "required": ["heading", "items"],
                            "properties": {
                                "heading": {"type": "string"},
                                "items": {"type": "array",
                                          "items": {"type": "string"}},
                            },
                        },
                    },
                    "steps": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "required": ["name"],
                            "properties": {
                                "name": {"type": "string"},
                                "actor": {"type": "string"},
                                "desc": {"type": "string"},
                                "sla": {"type": "string"},
                                "highlight": {"type": "boolean"},
                            },
                        },
                    },
                    "headers": {"type": "array", "items": {"type": "string"}},
                    "rows": {
                        "type": "array",
                        "items": {"type": "array", "items": {"type": "string"}},
                    },
                },
            },
        }
    },
}

SYSTEM = """\
Bạn là chuyên gia viết tài liệu hướng dẫn sử dụng phần mềm bằng tiếng Việt,
theo phong cách slide đào tạo nội bộ cho nhân viên cửa hàng.

Bạn nhận một loạt ảnh chụp màn hình theo đúng thứ tự thao tác. Nhiệm vụ:
đọc giao diện trong ảnh và viết nội dung slide hướng dẫn tương ứng.

QUY TẮC NỘI DUNG
- Viết cho người dùng cuối, không phải lập trình viên. Tránh từ kỹ thuật.
- Câu mệnh lệnh, ngắn gọn: "Chọn thiết bị", "Bấm Lưu".
- Mỗi callout tối đa 90 ký tự. Mỗi bullet tối đa 90 ký tự.
- Tối đa 6 callout hoặc 5 bullet một slide. Nhiều hơn thì tách slide.
- Chỉ mô tả những gì THỰC SỰ nhìn thấy trong ảnh. Không bịa nút, không bịa
  trường dữ liệu không có trên màn hình.
- Gọi đúng tên nút và nhãn trường như hiển thị trong ảnh.
- Dùng trường "note" cho cảnh báo hoặc lưu ý quan trọng, không lạm dụng.

QUY TẮC CẤU TRÚC
- Slide đầu tiên: layout "title".
- Chèn layout "section" trước mỗi nhóm thao tác lớn, "number" là "1", "2"...
- Ảnh có nhiều vùng cần chỉ dẫn theo thứ tự -> "screenshot_callout".
- Ảnh chỉ cần diễn giải chung -> "screenshot_bullets".
- Trường "image" PHẢI đặt đúng đường dẫn ảnh được cung cấp kèm mỗi ảnh.
- Slide không gắn ảnh (title, section, agenda, process, table) bỏ trống "image".
- Slide cuối: layout "table" tóm tắt các thao tác chính, rồi layout "title"
  để cảm ơn.
"""


def load_images(folder):
    files = sorted(f for f in os.listdir(folder) if f.lower().endswith(EXTS))
    if not files:
        sys.exit(f"Không tìm thấy ảnh nào trong: {folder}")
    out = []
    for fn in files:
        p = os.path.join(folder, fn)
        mt = mimetypes.guess_type(p)[0] or "image/png"
        with open(p, "rb") as f:
            b64 = base64.standard_b64encode(f.read()).decode()
        # đường dẫn ghi vào yaml - dùng dấu / cho gọn
        rel = os.path.join(folder, fn).replace("\\", "/")
        out.append((rel, mt, b64))
    return out


def build_content(folder, topic, extra=None):
    imgs = load_images(folder)
    print(f"Đọc {len(imgs)} ảnh từ {folder}/ ...")

    content = [{"type": "text", "text":
                f"Chủ đề tài liệu: {topic}\n\n"
                f"Dưới đây là {len(imgs)} ảnh chụp màn hình theo đúng thứ tự "
                f"thao tác. Mỗi ảnh kèm đường dẫn phải ghi vào trường image."}]

    if extra:
        content.append({"type": "text", "text":
                        f"Tài liệu nghiệp vụ tham khảo (ưu tiên dùng đúng "
                        f"thuật ngữ trong đây):\n\n{extra[:20000]}"})

    for path, mt, b64 in imgs:
        content.append({"type": "text", "text": f"\nẢnh - đường dẫn: {path}"})
        content.append({"type": "image", "source": {
            "type": "base64", "media_type": mt, "data": b64}})

    content.append({"type": "text", "text":
                    "Hãy sinh toàn bộ deck theo schema. Nhớ gán đúng đường dẫn "
                    "ảnh vào trường image của slide tương ứng."})

    client = anthropic.Anthropic()
    print(f"Gọi {MODEL} ...")
    with client.messages.stream(
        model=MODEL,
        max_tokens=32000,
        system=SYSTEM,
        thinking={"type": "adaptive"},
        output_config={
            "effort": "high",
            "format": {"type": "json_schema", "schema": SCHEMA},
        },
        messages=[{"role": "user", "content": content}],
    ) as stream:
        resp = stream.get_final_message()

    if resp.stop_reason == "refusal":
        sys.exit(f"Model từ chối: {resp.stop_details}")

    txt = "".join(b.text for b in resp.content if b.type == "text")
    u = resp.usage
    print(f"Token: vào {u.input_tokens}, ra {u.output_tokens}")
    return json.loads(txt)


def to_yaml(data, path):
    import yaml
    with open(path, "w", encoding="utf-8") as f:
        f.write("# File này do ai_fill.py sinh tự động.\n")
        f.write("# Hãy đọc lại và chỉnh tay trước khi render.\n\n")
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False,
                       default_flow_style=False, width=100)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", help="Thư mục chứa ảnh chụp màn hình")
    ap.add_argument("out", nargs="?", default="content_ai.yaml")
    ap.add_argument("--topic", required=True, help="Chủ đề tài liệu")
    ap.add_argument("--ref", help="File .md tài liệu nghiệp vụ tham khảo")
    a = ap.parse_args()

    extra = None
    if a.ref:
        with open(a.ref, encoding="utf-8") as f:
            extra = f.read()

    data = build_content(a.folder, a.topic, extra)
    to_yaml(data, a.out)
    print(f"[OK] {len(data.get('slides', []))} slide -> {a.out}")
    print(f"     Xem lại rồi chạy: python render.py {a.out} out/deck.pptx")


if __name__ == "__main__":
    main()
