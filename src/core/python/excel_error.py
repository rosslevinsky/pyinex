# Defines ExcelError (src/core/convert.cpp, start_conversion), given the texts of Excel's
# errors as a tuple.

def make(texts):
    class ExcelError:
        """An Excel error, such as #N/A.

        An error in a cell arrives in Python as an ExcelError, and a function returns one
        to put that error in its cell. ExcelError('#N/A') makes one; str(e) and e.text give
        its text.
        """
        __slots__ = ('_text',)
        __module__ = 'pyinex'
        __qualname__ = 'ExcelError'

        def __init__(self, text):
            # str.__str__ gives a plain str, so a subclass's own comparison is never used.
            plain = str.__str__(text) if isinstance(text, str) else None
            if plain not in texts:
                shown = repr(plain) if plain is not None else 'a ' + type(text).__name__
                raise ValueError(f'not an Excel error: {shown}; expected one of {", ".join(texts)}')
            self._text = plain

        @property
        def text(self):
            return self._text

        def __str__(self):
            return self._text

        def __repr__(self):
            return f'ExcelError({self._text!r})'

        def __eq__(self, other):
            if isinstance(other, ExcelError):
                return self._text == other._text
            return NotImplemented

        def __hash__(self):
            return hash(self._text)

    return ExcelError
