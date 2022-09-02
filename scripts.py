import os
import pandas as pd

from docxtpl import DocxTemplate
from docx2pdf import convert


def main(path_data: str,
         path_template: str,
         path_output: str) -> None:
    """

    :param path_data: path to table data
    :param path_template: path to .docx template
    :param path_output: path to the output folder
    """

    # data preprocessing
    data = pd.read_excel(path_data)

    for col in ['Фамилия', 'Имя', 'Отчетство']:
        data[col] = data[col].apply(lambda x: x.strip())

    # generating output directories
    for cat in data['Категория'].unique():
        sub_path = os.path.join(path_output, cat.replace('/', '_'))
        try:
            os.mkdir(sub_path)
        except FileExistsError:
            for file in os.listdir(sub_path):
                os.remove(os.path.join(sub_path, file))

    # generating files
    for ind in data.index:
        doc = DocxTemplate(path_template)
        content = {
            'uuid': data.loc[ind, 'ID участника'],
            'fio': ' '.join(data.loc[ind, ['Фамилия', 'Имя', 'Отчетство']].values),
            'specialization': 'Направление № 1',
            'degree': data.loc[ind, 'Категория'],
            'points_before': data.loc[ind, 'Баллы'],
            'points_after': data.loc[ind, 'Балл после апелляции'],
            'request': data.loc[ind, 'Апелляция'],
            'response': data.loc[ind, 'Ответ на апелляцию'],
        }
        doc.render(content)
        base_path = os.path.join(path_output,
                                 content['degree'].replace('/', '_'),
                                 str(content['uuid']))
        doc.save(f'{base_path}.docx')

    # transforming docx to pdf
    for i in os.listdir(path_output):
        sub_path = os.path.join(path_output, i)
        if os.path.isdir(sub_path):
            print(f'converting category {i}')
            convert(sub_path)
            for i in os.listdir(sub_path):
                if i.endswith('.docx'):
                    os.remove(os.path.join(sub_path, i))


if __name__ == 'main':
    path_data = "Данные для тестового задания.xlsx"
    path_template = "Образец.docx"
    path_output = "Готовые документы"

    main(path_data=path_data,
         path_template=path_template,
         path_output=path_output)
