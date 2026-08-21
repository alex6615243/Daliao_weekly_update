from docx import Document

import production_runner


app = production_runner.app
original_update_document = app.update_document


def update_document_without_stray_punctuation(template, output, target, data, config):
    original_update_document(template, output, target, data, config)
    document = Document(output)
    removed = False
    for paragraph in list(document.paragraphs):
        if paragraph.text.strip() == "。":
            element = paragraph._element
            element.getparent().remove(element)
            removed = True
    if removed:
        document.save(output)


app.update_document = update_document_without_stray_punctuation


if __name__ == "__main__":
    try:
        raise SystemExit(app.main())
    except Exception as exc:
        print(f"\n更新失敗：{exc}", file=app.sys.stderr)
        print("未覆寫任何原始檔。", file=app.sys.stderr)
        raise SystemExit(1)

