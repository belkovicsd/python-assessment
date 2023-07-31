from Task1_PPTX_report.main import get_type_mapping
from parameterized import parameterized


@parameterized.expand([
    ('title', 'generate_title_slide_report'),
    ('text', 'generate_text_slide_report'),
    ('list', 'generate_list_slide_report'),
    ('picture', 'generate_picture_slide_report'),
    ('plot', 'generate_plot_slide_report'),
])
def test_get_type_mapping(title, expected):
    assert expected in str(get_type_mapping(title))


def test_not_in_type_mapping():
    title = 'test'
    assert get_type_mapping(title) is None
