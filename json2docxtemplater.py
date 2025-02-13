"""
json2docxtemplater module, a layer on top of the docx package to support template replacements based on json
"""
import json
import re
from docx import Document
from docx.document import Document as Docx
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl

class Json2docx():
    ''' json2docxtemplater class '''
    DEFAULTCONFIG = {
        "debug" : True,
        "openmarker" : "#",
        "closemarker" : "#",
        "commentstyle" : "Comment",
        "headingstyle" : "Heading 1",
        "maxtablecolumns" : 9,
        "contentfolder" : "./",
        "templatefolder" : "./",
        "outputfolder" : "./"
    }

    def __init__(self, _config):
        self.config = self.DEFAULTCONFIG | _config
        self.table_list_regex_str = self.config["openmarker"] + r"([a-zA-Z0-9_]+)\(([a-zA-Z0-9_]+)" + (r",?([a-zA-Z0-9_]+)?"*(self.config["maxtablecolumns"]-1)) + r"\)" + self.config["closemarker"]
        self.table_list_regex = re.compile(self.table_list_regex_str)

    def debug(self, outstr):
        '''output only if debug is True'''
        if self.config["debug"]:
            print(outstr)

    def iter_block_items(self, parent):
        '''generator for docx paragraphs and tables, see https://github.com/python-openxml/python-docx/issues/276#issuecomment-199502885'''
        if isinstance(parent, Docx):
            parent_elm = parent.element.body
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        else:
            raise ValueError("unrecognized docx parent element")

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    def content_to_regex(self, contentdict):
        '''convert the content dictionary to a regex string matching either of the keys'''
        contentkeywords = [k for k,v in contentdict.items() if isinstance(v, type(""))]
        contenteregex = self.config["openmarker"] + "(?P<keyword>"+"|".join(contentkeywords)+")" + self.config["closemarker"]
        return re.compile(contenteregex)

    def replace_content(self, block, contentdict):
        '''execute replacements in block as defined in the contentdict'''
        if contentdict:
            contentregex = self.content_to_regex(contentdict)
            keywordmatch = contentregex.match(block.text)
            if keywordmatch:
                key = keywordmatch.group("keyword")
                self.debug(f"REPLACE {key} -> {contentdict[key]}")
                block.text = block.text.replace(self.config["openmarker"]+key+self.config["closemarker"], contentdict[key])

    def fill(self, contentfile, templatefile):
        '''execute replacements in the template file as defined in the content file'''
        doc = Document(self.config["templatefolder"] + templatefile)

        with open(self.config["contentfolder"] + contentfile, encoding="utf8") as file:
            content = json.load(file)
        localcontent = content
        tablelistconfig = []
        heading = "document start"

        for block in self.iter_block_items(doc):
            if isinstance(block, Paragraph):
                self.debug(f"P[{block.style.name}] {block.text}")
                tablelistmatch = self.table_list_regex.match(block.text)
                if tablelistmatch:
                    tablelistconfig = [s for s in tablelistmatch.groups() if s is not None]
                    block._element.getparent().remove(block._element)
                else:
                    if block.style.name == self.config["commentstyle"]:
                        self.debug("REMOVE")
                        block._element.getparent().remove(block._element)
                    else:
                        if len(tablelistconfig) == 2: # list name + a single field
                            if block.style.name == self.config["headingstyle"]:
                                raise ValueError(f"list definition is followed by heading in the end of heading '{heading}', please add a newline in between in the template!")
                            if tablelistconfig[0] in localcontent.keys():
                                recordid = 0
                                for record in localcontent[tablelistconfig[0]]:
                                    recordid += 1
                                    if "bullet" in tablelistconfig[0]:
                                        listitemtext = chr(9679) + " " + record
                                    else:
                                        listitemtext = str(recordid) + ". " + record
                                    self.debug(f"new list item, FILL {listitemtext}")
                                    block.insert_paragraph_before(text=listitemtext)
                            else:
                                raise ValueError(f"no content in content file, '{contentfile}', for template list definition '{tablelistconfig[0]}' under heading '{heading}'")
                        elif block.style.name == self.config["headingstyle"]:
                            heading = block.text
                            headingmatches = [k for k in content.keys() if heading.lower().startswith(k)]
                            if len(headingmatches) == 1:
                                self.debug(f"H[---] {headingmatches[0]}")
                                localcontent = content[headingmatches[0]]
                            elif len(headingmatches) == 0:
                                self.debug(f"H[---] ***unrecognized*** no json content key prefix-matches '{heading.lower()}', falling back to document root")
                                localcontent = content
                                heading += " (no content match)"
                            else:
                                raise ValueError(f"more prefix matches for heading, '{heading}', in content file, '{contentfile}'")

                        self.replace_content(block, localcontent)
                        tablelistconfig = []
            elif isinstance(block, Table):
                self.debug(f"T[{block.style.name if block.style is not None else 'no style'}] config={','.join(tablelistconfig)}")
                if len(tablelistconfig) > 1:
                    if tablelistconfig[0] in localcontent.keys():
                        for record in localcontent[tablelistconfig[0]]:
                            self.debug("new row")
                            newrow = block.add_row()
                            for cellidx in range(1, len(tablelistconfig)):
                                self.debug(f"FILL {cellidx} {record[tablelistconfig[cellidx]]}")
                                newrow.cells[cellidx-1].text = record[tablelistconfig[cellidx]]
                    else:
                        raise ValueError(f"no content in content file, '{contentfile}', for tamplate table definition '{tablelistconfig[0]}' under heading '{heading}'")
                else:
                    self.debug("no configuration for template table, skipping")
                tablelistconfig = []
            else:
                self.debug(f"?[{block.style.name}] class={block.__class__.__name__}")
        self.debug(f"OUTPUT: {self.config['outputfolder'] + contentfile + templatefile}")
        doc.save(self.config["outputfolder"] + contentfile + "." + templatefile)
