import docx

WORD_NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
INS_TAG = f"{{{WORD_NAMESPACE['w']}}}ins"
DEL_TAG = f"{{{WORD_NAMESPACE['w']}}}del"


def accept_all_revisions_in_doc(doc):
    """
    Accept all revisions (track changes) in the docx document by searching for `w:ins`
    (inserted elements) and `w:del` (deleted elements) in the entire XML tree,
    including headers and footers.

    :param doc: A docx document object to process.
    """
    # main body only
    _process_element(doc.element)
    # each seaction separately (e.g. heaers and footers)
    for section in doc.sections:
        _process_element(section.header._element)
        _process_element(section.footer._element)

    _process_footnotes_endnotes(doc)
    _process_comments(doc)
    _process_textboxes(doc)

def _process_element(element):
    """Process any XML element to accept insertions and remove deletions."""
    _accept_all_insertions(element)
    _remove_all_deletions(element)

def _accept_all_insertions(element):
    """Accept all inserted content in the document by keeping `w:ins` elements."""
    for ins in element.findall(f".//{INS_TAG}"):
        parent = ins.getparent()
        for child in ins:
            parent.insert(parent.index(ins), child)
        parent.remove(ins)

def _remove_all_deletions(element):
    """Remove all deleted content in the document by removing `w:del` elements."""
    for deletion in element.findall(f".//{DEL_TAG}"):
        deletion.getparent().remove(deletion)

def _process_footnotes_endnotes(doc):
    """Process footnotes and endnotes in the document."""
    footnotes_part = doc.part.related_parts.get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes}')
    endnotes_part = doc.part.related_parts.get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes}')

    if footnotes_part:
        footnotes_xml = footnotes_part.element
        _process_element(footnotes_xml)

    if endnotes_part:
        endnotes_xml = endnotes_part.element
        _process_element(endnotes_xml)

def _process_comments(doc):
    """Process comments in the document."""
    comments_part = doc.part.related_parts.get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments}')
    if comments_part:
        comments_xml = comments_part.element
        _process_element(comments_xml)

def _process_textboxes(doc):
    """Process textboxes and shapes in the document."""
    for shape in doc.element.findall(f'.//{{{WORD_NAMESPACE["w"]}}}textbox'):
        _process_element(shape)