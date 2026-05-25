# UEH MBA Thesis — Markdown + Pandoc pipeline

Đề án Thạc sĩ MBA (Trường Kinh doanh, UEH) viết bằng Markdown, build sang Word `.docx` chuẩn UEH thông qua Pandoc + python-docx.

## Vì sao không viết thẳng trong Word?

- Nguồn là plain text → **diff/merge bằng git**, revert nếu cần.
- AI agent (Claude Code) có thể đọc/sửa/review trực tiếp trên `.md`.
- Tách bạch **nội dung** (Markdown) khỏi **trình bày** (template docx).

## Cấu trúc

```
.
├── Makefile                  # build pipeline
├── metadata.yaml             # title, author, bib, csl
├── references.bib            # nguồn trích dẫn (BibTeX)
├── apa.csl                   # style trích dẫn APA
├── build_reference.py        # sinh reference-ueh.docx (styles UEH)
├── make_figures.py           # sinh hình demo bằng matplotlib
├── post_process.py           # chèn SEQ field cho caption Hình/Bảng
├── figures/                  # hình minh hoạ (input)
│   ├── hinh1.png
│   └── hinh2.png
├── outputs/                  # gitignored — chứa luanvan.docx, reference-ueh.docx sau khi build
└── chapters/                 # nội dung — sửa ở đây
    ├── 00-bia.md
    ├── 01-bia-phu.md
    ├── 02-loi-cam-doan.md
    ├── 03-loi-cam-on.md
    ├── 04-muc-luc.md
    ├── 05-chuong-1-tong-quan.md
    ├── 06-chuong-2-co-so-ly-luan.md
    ├── 07-chuong-3-phuong-phap.md
    ├── 08-chuong-4-thuc-trang.md
    ├── 09-chuong-5-giai-phap.md
    ├── 10-ket-luan.md
    └── 11-tai-lieu-tham-khao.md
```

Cấu trúc 5 chương + Kết luận + TLTK theo **Hướng dẫn Cấu trúc nội dung Đề án Thạc sĩ MBA** của Khoa Quản trị, Trường Kinh doanh — UEH (2026).

## Yêu cầu

- [Pandoc](https://pandoc.org/) ≥ 3.0
- Python ≥ 3.10
- `make`

## Setup lần đầu

```bash
python3 -m venv venv
./venv/bin/pip install python-docx matplotlib
```

## Build

```bash
make            # → luanvan.docx
make open       # build + mở trong Word
make clean      # xoá output
make figures    # regenerate hình demo
make reference  # rebuild template style UEH
```

Mở file trong Word, nhấn **F9** trên các trang Mục lục / Danh mục hình / Danh mục bảng để cập nhật.

## Chuẩn UEH đã thiết lập

| Mục              | Quy chuẩn                                                   |
|------------------|-------------------------------------------------------------|
| Khổ giấy         | A4                                                          |
| Lề               | Trên 3 cm, dưới 3 cm, trái 3.5 cm, phải 2 cm                |
| Font             | Times New Roman 13pt, màu đen                                |
| Giãn dòng        | 1.5                                                         |
| Thụt đầu dòng    | 1 cm                                                        |
| Căn lề           | Justify                                                     |
| Heading 1        | TNR 14pt, **đậm**, IN HOA, căn giữa                         |
| Heading 2        | TNR 13pt, **đậm**, căn trái                                 |
| Heading 3        | TNR 13pt, ***đậm nghiêng***, căn trái                       |
| Bảng             | Border đen 0.75pt, header xám `#D9D9D9` đậm, width 100%     |
| Caption          | Auto-number bằng SEQ field, reset theo chương               |
| Trích dẫn        | APA (citeproc + `apa.csl`)                                  |
| Số trang         | Footer căn giữa                                             |

## Workflow đề xuất

1. Sửa nội dung trong `chapters/*.md`.
2. `make open` để xem kết quả.
3. Commit theo từng section/chương cho dễ revert.
4. Branch theo vòng phản hồi GVHD: `gvhd-v1`, `gvhd-v2`…
5. Gần nộp: đóng băng `.md`, xuất `.docx` lần cuối, từ đó chỉ sửa trên Word theo Track Changes của GVHD.

## Tham khảo

- Hướng dẫn cấu trúc Đề án Thạc sĩ MBA — Khoa Quản trị, Trường Kinh doanh, UEH (2026).
