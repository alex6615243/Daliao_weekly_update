from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import ship_runner


app = ship_runner.app


def replace_text_preserving_drawings(paragraph, text: str) -> None:
    text_nodes = paragraph._element.xpath(".//w:t")
    if text_nodes:
        text_nodes[0].text = text
        for node in text_nodes[1:]:
            node.text = ""
        return
    run = OxmlElement("w:r")
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    paragraph._element.append(run)


app.replace_paragraph_text = replace_text_preserving_drawings


if __name__ == "__main__":
    try:
        raise SystemExit(app.main())
    except Exception as exc:
        print(f"\n更新失敗：{exc}", file=app.sys.stderr)
        print("未覆寫任何原始檔。", file=app.sys.stderr)
        raise SystemExit(1)

