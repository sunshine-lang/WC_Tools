"""准备流程共用的标签、占位符与可见文本长度规则。"""
import re


# %i、%f 后紧跟字母时不是参数，如德语 100%ige。
PARAM_PATTERN = re.compile(r'\{\w*\}|%\([^)]+\)[sdif]|%(?:\d+\$)?(?:[sd]|[if](?![^\W\d_]))')
TAG_PATTERN = re.compile(r'</?[A-Za-z][^<>]*>')


def display_length(value):
    """标签与换行不计长；空格、标点和参数按字面计入。"""
    return len(re.sub(r'\\n|\n', '', TAG_PATTERN.sub('', value)))
