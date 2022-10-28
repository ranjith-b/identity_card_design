from docxtpl import DocxTemplate
import pandas as pd
import numpy as np
from docx2pdf import convert


class Visitor:
    def __init__(self, name, unit, accompanying_adult, children_1, children_2, children_3, contact_number):
        self.name = name
        self.unit = unit
        self.accompanying_adult = accompanying_adult
        self.children_1 = children_1
        self.children_2 = children_2
        self.children_3 = children_3
        self.contact_number = contact_number


def create_tags(visitors):
    for visitor in visitors:
        doc = DocxTemplate("template.docx")
        context = {
            'name': visitor.name,
            'unit': visitor.unit,
            'accompanying_adult': visitor.accompanying_adult,
            'children_1': visitor.children_1,
            'children_2': visitor.children_2,
            'children_3': visitor.children_3,
            'contact_number': visitor.contact_number
        }
        doc.render(context)
        # doc.save(f"{visitor.name[:5]}.docx")
        doc.save('C:\CGI_projects\identity_card_design\\badges_in_word_format\\' +
                 f"{visitor.name[:5]}.docx")
        convert('C:\CGI_projects\identity_card_design\\badges_in_word_format\\' +
                f"{visitor.name[:5]}.docx", 'C:\CGI_projects\identity_card_design\\badges_in_pdf_format\\' + f"{visitor.name[:5]}.pdf")
        #convert(f"{visitor.name[:5]}.docx", f"{visitor.name[:5]}.pdf")


visitor_data = pd.read_excel('data_ids.xlsx', index_col=0)
visitor_data = visitor_data.replace(np.nan, '-', regex=True)

# print(visitor_data['name'])

for index, row in visitor_data.iterrows():
    visitor_data = [row.name,
                    row.unit,
                    row.accompanying_adult,
                    row.children_1,
                    row.children_2,
                    row.children_3,
                    row.contact_number]

    visitors = [Visitor(*visitor_data)]
    create_tags(visitors)
