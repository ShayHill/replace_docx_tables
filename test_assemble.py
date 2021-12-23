#!/usr/bin/env python3
# last modified: 211222 19:43:15
"""Test assembling an entire docx from some data and a template.

:author: Shay Hill
:created: 2021-12-22
"""

from pathlib import Path
from typing import List, Tuple, Union

from lxml import etree

from docx2python import docx2python
from docx2python.attribute_register import Tags
from docx2python.utilities import replace_root_text


TEMPLATE = "template.docx"
OUTPUT_FILE = "template_filled_in.docx"


def has_text(root: etree._Element, text: str) -> bool:
    """Does any descendent element of :root: contain :text:?

    :param root: an etree Element
    :param text: text to search for in root descendents
    """
    return any(text in (x.text or "") for x in root.iter())


def _find_text(root: etree._Element, text: str, tag: str) -> etree._Element:
    """Find next descendent element that contains :text:

    :param root: etree.Element that presumably contains a descendent element with :text:
    :param text: text in table that will identify elem
    :param _tag: type of element sought
    :returns: elem.tag == :tag: and has_text(:text:)
    :raises: ValueError if no matching element can be found
    """
    try:
        return next(x for x in root.iter() if x.tag == tag and has_text(x, text))
    except StopIteration:
        raise ValueError(f"{text} not found in {tag} element in {root}")


def get_table_elements(
    root: etree._Element, marker_text: str
) -> Tuple[etree._Element, etree._Element, etree._Element]:
    """Find the (first) table, table_row, and table_cell that contain :text:

    :param root: etree.Element that presumably contains a descendent table with :text:
    :param marker_text: text in table that will identify iter
    :returns: the first table in root that contains :text:,
    the first table_row in that table that containss :text:,
    the first table_cell in that table_row that contains :text:
    :raises: ValueError if text not found in a table
    """
    table = _find_text(root, marker_text, Tags.TABLE)
    table_row = _find_text(table, marker_text, Tags.TABLE_ROW)
    table_cell = _find_text(table_row, marker_text, Tags.TABLE_CELL)
    table_row.remove(table_cell)
    table.remove(table_row)
    return table, table_row, table_cell


def insert_table_rows(
    template: Union[Path, str],
    marker_text: str,
    rows: List[List[str]],
    output_path: Union[Path, str],
) -> None:
    """Insert values in rows as table values

    :param template: path to template *docx file
    :param marker_text: text in template destination table cell
    :param rows: rows of cells of text to be inserted into template
    :param output_path: path to output for updated template
    """
    template_reader = docx2python(template).docx_reader
    template_root = template_reader.file_of_type("officeDocument").root_element
    table, table_row, table_cell = get_table_elements(template_root, marker_text)
    for row in rows:
        new_row = etree.fromstring(etree.tostring(table_row))
        for cell in row:
            new_cell = etree.fromstring(etree.tostring(table_cell))
            replace_root_text(new_cell, marker_text, cell)
            new_row.append(new_cell)
        table.append(new_row)
    template_reader.save(output_path)


def test_insert_table_rows():
    insert_table_rows(
        TEMPLATE, "CELL TEXT", [["1", "2", "3"], ["4", "5", "6"]], OUTPUT_FILE
    )

    assert docx2python(OUTPUT_FILE).document_runs == [
        [[[["HEADING PARAGRAPH"]]]],
        [[[["1"]], [["2"]], [["3"]]], [[["4"]], [["5"]], [["6"]]]],
        [[[[]]]],
    ]
