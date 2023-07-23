import logging

from pptx import Presentation
import pandas
from pptx.util import Inches

CONTENT_KEY = 'content'
FILE_NAME = 'example_output.pptx'


class UnknownTypeException(Exception):
    pass


def read_config():
    clear_presentation()
    file = pandas.read_json('sample.json')

    type_mapping = {
        'title': generate_title_slide_report,
        'text': generate_text_slide_report,
        'list': generate_list_slide_report,
        'picture': generate_picture_slide_report,
        'plot': generate_plot_slide_report
    }

    for index, row in file.iterrows():

        first_element = row.iloc[0]
        type_value = first_element['type']

        process_object = type_mapping.get(type_value)

        try:
            if process_object is None:
                raise UnknownTypeException("Invalid 'type' value:")

            else:
                logging.debug(f"read_config - Processing record, Type: '{type_value}'")
                process_object(first_element)

        except UnknownTypeException as e:
            logging.warning(f"read_config - {e} '{type_value}', Index: {index}")


def clear_presentation():
    Presentation().save(FILE_NAME)


def create_paragraph_for_each_element(elements, slide):
    shape = slide.shapes.placeholders[1]
    text_frame = shape.text_frame
    for element in elements:
        paragraph = text_frame.add_paragraph()
        paragraph.text = element['text']
        paragraph.level = element['level']


def create_presentation(layout, data):
    presentation = Presentation(FILE_NAME)
    slide_layout = presentation.slide_layouts[layout]
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = pandas.Series(data)['title']

    return presentation, slide


def generate_title_slide_report(data):
    presentation, slide = create_presentation(0, data)

    subtitle = slide.placeholders[1]
    subtitle.text = pandas.Series(data)[CONTENT_KEY]

    presentation.save(FILE_NAME)


def generate_text_slide_report(data):
    presentation, slide = create_presentation(5, data)

    left = top = Inches(1.2)
    width = height = Inches(5)
    text_box = slide.shapes.add_textbox(left, top, width, height)
    text_box.text_frame.text = pandas.Series(data)[CONTENT_KEY]

    presentation.save(FILE_NAME)


def generate_list_slide_report(data):
    presentation, slide = create_presentation(1, data)

    elements = pandas.Series(data)[CONTENT_KEY]
    create_paragraph_for_each_element(elements, slide)

    presentation.save(FILE_NAME)



def generate_picture_slide_report(data):
    print(data)


def generate_plot_slide_report(data):
    print(data)


read_config()
