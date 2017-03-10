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

    def __repr__(self):
        suffix = ["{}, {}".format(self.family, self.size)]
        if self.bold:
            suffix.append('b')
        if self.italic:
            suffix.append('i')
        if self.underline:
            suffix.append('u')

        return "<Font '{}'>".format(' '.join(suffix))

if __name__ == '__main__':
    f = Font(bold=True, italic=True)
    print(f)
