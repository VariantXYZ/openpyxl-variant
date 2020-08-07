# Copyright (c) 2010-2020 openpyxl

from openpyxl.cell.text import Text

from openpyxl.xml.functions import iterparse
from openpyxl.xml.constants import SHEET_MAIN_NS


def read_string_table(xml_source):
    """Read in all shared strings in the table"""

    strings = []
    STRING_TAG = '{%s}si' % SHEET_MAIN_NS

    for _, node in iterparse(xml_source):
        if node.tag == STRING_TAG:
            original_text = Text.from_tree(node)

            if len(original_text.formatted) == 0:
                text = original_text.content
            else:
                text = ""
                for x in original_text.formatted:
                    txt = x.t # Will handle InlineFont and RichText
                    if x.rPr:
                        if x.rPr.b == True:
                            txt = f"<b>{txt}</b>"
                        if x.rPr.i == True:
                            txt = f"<i>{txt}</i>"
                    text += txt

            text = text.replace('x005F_', '')
            node.clear()
            strings.append(text)

    return strings
