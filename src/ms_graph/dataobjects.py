from dataclasses import dataclass, field

from typing import List


@dataclass
class TextColumn:
    allowMultipleLines: bool = True
    appendChangesToExistingText: bool = False
    linesForEditing: int = 1
    maxLength: int = 255
    textType: str = "plain"


@dataclass
class DateTimeColumn:
    displayAs: str = 'default'
    format: str = 'dateTime'


@dataclass
class ColumnDefinition:
    name: str
    displayName: str
    description: str = ''
    text: TextColumn = None
    dateTime: DateTimeColumn = None
    hidden: bool = False
    required: bool = False
    readOnly: bool = False


@dataclass
class ListInfo:
    contentTypesEnabled: bool = False
    hidden: bool = False
    template: str = 'genericList'


@dataclass
class SharepointList:
    displayName: str
    columns: List[ColumnDefinition]
    list: ListInfo = field(default_factory=ListInfo)


def get_col_definition(type, **kwargs):
    col_def = None
    if type == 'text':
        col_def = TextColumn(**kwargs)
    elif type == 'dateTime':
        col_def = DateTimeColumn(format='dateTime')
    elif type == 'date':
        col_def = DateTimeColumn(format='dateOnly')
    else:
        raise ValueError(f'Unsupported column type: {type}')
    return col_def


def get_col_def_name(type):
    if type in ['date', 'dateTime']:
        return 'dateTime'
    else:
        return type
