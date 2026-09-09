# -*- coding: utf-8 -*-
"""
Kiểm tra ảnh chụp màn hình trước khi render.

    python check_shots.py                  # kiểm tra content.yaml
    python check_shots.py content_ai.yaml
    python check_shots.py --todo           # in danh sách ảnh cần chụp

Trả lời đúng câu hỏi "thay hình sao cho chính xác":
  - Slide nào cần ảnh nào, tiêu đề slide là gì, callout nói về cái gì
  - Ảnh nào còn thiếu, ảnh nào thừa (có trong shots/ nhưng không slide nào dùng)
  - Ảnh nào độ phân giải thấp, tỉ lệ lệch so với khung, hoặc quá nặng
"""
import os, sys, io, json

for _s in ("stdout", "stderr"):
    _st = getattr(sys, _s, None)
    if _st and hasattr(_st, "buffer"):
        setattr(sys, _s, io.TextIOWrapper(_st.buffer, encoding="utf-8",
                                          errors="replace", line_buffering=True))

# Khung ảnh thực tế trong slide (pt) - khớp với render.py
FRAME = {
    "screenshot_callout": (530, 348),
    "screenshot_bullets": (510, 348),
}
MIN_W = 900          # dưới mức này in ra sẽ mờ
MAX_MB = 4.0


def load(path):
    with open(path, encoding="utf-8") as f:
        t = f.read()
    if path.lower().endswith((".yaml", ".yml")):
        import yaml
        return yaml.safe_load(t)
    return json.loads(t)


def describe(s):
    """Tóm tắt slide cần ảnh gì - để chụp cho đúng màn hình."""
    out = []
    for co in s.get("callouts", []):
        out.append(co.get("text", "") if isinstance(co, dict) else str(co))
    out += [str(b) for b in s.get("bullets", [])]
    return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    todo_only = "--todo" in sys.argv
    src = args[0] if args else "content.yaml"
    data = load(src)
    slides = data.get("slides", [])

    need, missing, ok, warn = [], [], [], []

    for i, s in enumerate(slides, start=1):
        img = s.get("image")
        if not img:
            continue
        lay = s.get("layout", "")
        need.append(img)
        rec = (i, img, s.get("title", ""), lay, describe(s))

        if not os.path.exists(img):
            missing.append(rec)
            continue

        # --- ảnh có thật: soi chất lượng ---
        from PIL import Image
        try:
            w, h = Image.open(img).size
        except Exception as e:
            warn.append((i, img, f"không mở được ảnh: {e}"))
            continue

        mb = os.path.getsize(img) / 1048576
        msgs = []
        if w < MIN_W:
            msgs.append(f"rộng {w}px - dưới {MIN_W}px sẽ mờ khi chiếu")
        if mb > MAX_MB:
            msgs.append(f"nặng {mb:.1f}MB - nên nén lại")

        fw, fh = FRAME.get(lay, (530, 348))
        r_img, r_frame = w / h, fw / fh
        if r_img > r_frame * 1.9:
            msgs.append(f"ảnh quá ngang ({w}x{h}) - vào khung sẽ bị nhỏ, "
                        f"nên cắt bớt bề ngang")
        elif r_img < r_frame * 0.55:
            msgs.append(f"ảnh quá cao ({w}x{h}) - vào khung sẽ bị nhỏ, "
                        f"nên cắt bớt chiều cao")

        (warn if msgs else ok).append((i, img, msgs, f"{w}x{h}", f"{mb:.1f}MB"))

    # ---------------- báo cáo ----------------
    if missing:
        print("=" * 68)
        print(f"CẦN CHỤP {len(missing)} ẢNH")
        print("=" * 68)
        for i, img, title, lay, hints in missing:
            print(f"\n[slide {i}]  {img}")
            print(f"  Tiêu đề : {title}")
            print(f"  Cần thấy được trên màn hình:")
            for h in hints[:6]:
                print(f"     - {h}")
        print()

    if todo_only:
        return

    if warn:
        print("=" * 68)
        print(f"CÓ ẢNH NHƯNG NÊN XEM LẠI ({len(warn)})")
        print("=" * 68)
        for w_ in warn:
            if len(w_) == 3:
                print(f"  slide {w_[0]:>2}  {w_[1]}\n     ! {w_[2]}")
            else:
                i, img, msgs, dim, mb = w_
                print(f"  slide {i:>2}  {img}  ({dim}, {mb})")
                for m in msgs:
                    print(f"     ! {m}")
        print()

    if ok:
        print(f"ẢNH ĐẠT: {len(ok)}")
        for i, img, _, dim, mb in ok:
            print(f"  slide {i:>2}  {img}  ({dim}, {mb})")
        print()

    # --- ảnh nằm trong shots/ nhưng không slide nào dùng ---
    folders = {os.path.dirname(p) for p in need if os.path.dirname(p)}
    orphan = []
    for d in sorted(folders):
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if f.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
                p = os.path.join(d, f).replace("\\", "/")
                if p not in need:
                    orphan.append(p)
    if orphan:
        print("=" * 68)
        print("ẢNH THỪA - có trong thư mục nhưng không slide nào dùng")
        print("=" * 68)
        for p in orphan:
            print(f"  {p}")
        print("  -> sai tên file? So lại với danh sách bên trên.")
        print()

    print("-" * 68)
    print(f"Tổng: {len(need)} ảnh cần | {len(ok)} đạt | "
          f"{len(warn)} cần xem lại | {len(missing)} chưa có")
    if not missing and not warn:
        print("Sẵn sàng render.")


if __name__ == "__main__":
    main()
