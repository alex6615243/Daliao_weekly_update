from zipfile import ZipFile

from docx import Document

import release_runner


app = release_runner.app


def validate_docx(path):
    with ZipFile(path) as package:
        if package.testzip() is not None:
            raise RuntimeError("輸出 Word 檔案結構損壞。")
    texts = [paragraph.text for paragraph in Document(path).paragraphs]
    if not any("《代禱事項》" in text for text in texts):
        raise RuntimeError("輸出缺少《代禱事項》。")
    if not any("《本週報告事項》" in text for text in texts):
        raise RuntimeError("輸出缺少《本週報告事項》。")


app.validate_docx = validate_docx


if __name__ == "__main__":
    try:
        raise SystemExit(app.main())
    except Exception as exc:
        print(f"\n更新失敗：{exc}", file=app.sys.stderr)
        print("未覆寫任何原始檔。", file=app.sys.stderr)
        raise SystemExit(1)

