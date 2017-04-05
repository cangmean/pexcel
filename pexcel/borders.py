# coding=utf-8


class Border(object):
    style_map = {
        'dashDot',
        'dashDotDot',
        'dashed',
        'dotted',
        'double',
        'hair',
        'medium',
        'mediumDashDot',
        'mediumDashDotDot',
        'mediumDashed',
        'slantDashDot',
        'thick',
        'thin',
    }

    def __init__(self, border_style=None):
        self.border_style = border_style

    @property
    def is_default(self):
        return self.hash_key == Borders().hash_key

    @property
    def hash_key(self):
        return hash(self.border_style)

    def __eq__(self, other):
        return self.hash_key == other.hash_key

    def __hash__(self):
        return self.hash_key

    def __repr__(self):
        return "<Border '{}'>".format(self.border_style or 'default')


class Borders(object):

    def __init__(self, left=None, right=None, top=None, bottom=None):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom

    @property
    def is_default(self):
        return self.hash_key == Borders().hash_key

    @property
    def hash_key(self):
        return hash((self.left, self.right, self.top, self.bottom))

    def __eq__(self, other):
        return self.hash_key == other.hash_key

    def __hash__(self):
        return self.hash_key

    def __repr__(self):
        suffix = []
        for key in ('left', 'right', 'top', 'bottom'):
            border = getattr(self, key, None)
            if border:
                suffix.append('{}: {}'.format(key, border))
        return "<Borders '{}'>".format(' '.join(suffix) or 'default')
