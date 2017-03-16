# coding=utf-8


class Font(object):

    def __init__(self, bold=False, italic=False, underline=False,
                 strikethrough=False, family='Calibri', size=11, color=None):
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strikethrough = strikethrough
        self.family = family
        self.size = size
        self.color = color

    def __eq__(self, other):
        return (self.bold, self.italic, self.underline,
                self.strikethrough, self.family, self.size,
                self.color) == \
                (other.bold, other.italic, other.underline,
                 other.strikethrough, other.family, other.size,
                 other.color)

    def __hash__(self):
        # 通过__eq__ 和 __hash__ 指定字典的key相同
        return hash((self.bold, self.italic, self.underline,
                     self.strikethrough, self.family, self.size,
                     self.color))

    def __repr__(self):
        suffix = ["{}, {}".format(self.family, self.size)]
        if self.bold:
            suffix.append('b')
        if self.italic:
            suffix.append('i')
        if self.underline:
            suffix.append('u')

        return "<Font '{}'>".format(' '.join(suffix))
