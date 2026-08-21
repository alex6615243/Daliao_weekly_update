import re

from docx.oxml import OxmlElement

import current_runner


app = current_runner.app


def replace_text_with_template_run(paragraph, text: str) -> None:
    existing = paragraph.text
    if re.search(r"\d+期", existing) and re.search(r"\d+期", text):
        large_nodes = paragraph._element.xpath(".//w:r[w:rPr/w:sz[@w:val='70']]/w:t")
        issue = re.search(r"(\d+期)", text).group(1)
        if large_nodes:
            for index, node in enumerate(large_nodes):
                node.text = issue[index] if index < len(issue) else ""
            if len(issue) > len(large_nodes):
                large_nodes[0].text = issue
                for node in large_nodes[1:]:
                    node.text = ""
            return

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


app.replace_paragraph_text = replace_text_with_template_run


if __name__ == "__main__":
    try:
        raise SystemExit(app.main())
    except Exception as exc:
        print(f"\n更新失敗：{exc}", file=app.sys.stderr)
        print("未覆寫任何原始檔。", file=app.sys.stderr)
        raise SystemExit(1)

