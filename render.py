from docxtpl import DocxTemplate
import pandas as pd


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
        doc.save(f"generated_{visitor.name}.docx")


visitor_data = pd.read_excel('data_ids.xlsx', index_col=0)
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
