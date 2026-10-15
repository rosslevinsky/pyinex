"""How PyCall finds a script in real Excel (docs/architecture.md, "PyCall" and "Scripts"):
absolute and relative paths, the workbook's folder as Excel gives it, and which module a
path is.

Helper modules are cached by their plain names for the whole session, so each test's
helpers have names no other test uses.
"""

from conftest import E, pycall

UNICODE = "Bücher ✓ 😀"


def returns(text):
    return f"def f():\n    return {text!r}\n"


def test_an_absolute_path(book, tmp_path):
    script = tmp_path / "elsewhere" / "absolute.py"
    script.parent.mkdir()
    script.write_text(returns("by its full path"), encoding="utf-8")
    assert book.formula("A1", pycall(str(script), "f")) == "by its full path"


def test_a_relative_path_is_from_the_workbooks_folder(book):
    book.script("relative.py", returns("beside the workbook"))
    assert book.formula("A1", pycall("relative.py", "f")) == "beside the workbook"


def test_a_relative_path_into_a_subfolder(book):
    (book.folder / "sub").mkdir()
    (book.folder / "sub" / "deeper.py").write_text(returns("in a subfolder"), encoding="utf-8")
    assert book.formula("A1", pycall(r"sub\deeper.py", "f")) == "in a subfolder"


def test_a_workbook_whose_only_sheet_has_its_name(workbook, tmp_path):
    # Excel's CELL("filename") then gives C:\folder\Named.xlsx, with no [Named.xlsx]Named.
    named = workbook(tmp_path / "Named.xlsx", sheet="Named")
    assert named.formula("B1", '=CELL("filename",A1)') == str(tmp_path / "Named.xlsx")
    named.script("alone.py", returns("found without a sheet"))
    assert named.formula("A1", pycall("alone.py", "f")) == "found without a sheet"


def test_a_folder_with_brackets(workbook, tmp_path):
    bracketed = workbook(tmp_path / "data[old]" / "book.xlsx")
    bracketed.script("bracketed.py", returns("in a bracketed folder"))
    assert bracketed.formula("A1", pycall("bracketed.py", "f")) == "in a bracketed folder"


def test_a_folder_with_brackets_and_a_sheet_named_like_the_workbook(workbook, tmp_path):
    both = workbook(tmp_path / "data[old]" / "Same.xlsx", sheet="Same")
    both.script("both.py", returns("both at once"))
    assert both.formula("A1", pycall("both.py", "f")) == "both at once"


def test_an_unsaved_workbook_refuses_a_relative_path(workbook, cleared, tmp_path):
    unsaved = workbook()
    assert unsaved.formula("A1", pycall("anything.py", "f")) == E("#VALUE!")
    said = "Pyinex: cannot find the script anything.py: a relative path is relative to the workbook's folder, and this workbook has not been saved"
    window = cleared(said)
    assert window and said in window.text
    script = tmp_path / "absolute_from_unsaved.py"
    script.write_text(returns("an absolute path still works"), encoding="utf-8")
    assert unsaved.formula("A2", pycall(str(script), "f")) == "an absolute path still works"


def test_two_workbooks_share_one_script_and_its_module_state(workbook, tmp_path):
    first = workbook(tmp_path / "first" / "first.xlsx")
    second = workbook(tmp_path / "second" / "second.xlsx")
    first.script("shared.py", "n = 0\n\ndef count():\n    global n\n    n += 1\n    return n\n")
    shared = first.folder / "shared.py"
    assert first.formula("A1", pycall("shared.py", "count")) == 1.0
    assert second.formula("A1", pycall(str(shared), "count")) == 2.0
    # The same file named another way, in another case, is the same module.
    assert second.formula("A2", pycall(str(shared).upper(), "count")) == 3.0


def test_same_named_scripts_in_two_folders_are_two_modules(workbook, tmp_path):
    # Excel opens only one workbook of a name at a time, so the workbooks' names differ.
    first = workbook(tmp_path / "one" / "first.xlsx")
    second = workbook(tmp_path / "two" / "second.xlsx")
    first.script("model.py", returns("the first model"))
    second.script("model.py", returns("the second model"))
    assert first.formula("A1", pycall("model.py", "f")) == "the first model"
    assert second.formula("A1", pycall("model.py", "f")) == "the second model"


def test_unicode_and_spaces_in_every_name(workbook, tmp_path):
    folder = tmp_path / UNICODE
    book = workbook(folder / f"{UNICODE}.xlsx", sheet="Лист ✓ 😀")
    (folder / "hilfe_ü_paths.py").write_text("WERT = 'aus der Hilfe'\n", encoding="utf-8")
    book.script(
        f"Skript {UNICODE}.py",
        "import pyinex\nimport hilfe_ü_paths\n\n"
        "def f():\n    return [hilfe_ü_paths.WERT, pyinex.sheet(), pyinex.caller(full=True), __file__]\n",
    )
    book.formula("A1", pycall(f"Skript {UNICODE}.py", "f"))
    sheet = f"[{UNICODE}.xlsx]Лист ✓ 😀"
    assert book.values("A1:D1") == [["aus der Hilfe", sheet, sheet + "!A1", str(folder / f"Skript {UNICODE}.py")]]


def test_a_helper_beside_the_script(book):
    (book.folder / "helper_paths.py").write_text("def greet():\n    return 'from the helper'\n", encoding="utf-8")
    book.script("uses_helper.py", "import helper_paths\n\ndef f():\n    return helper_paths.greet()\n")
    assert book.formula("A1", pycall("uses_helper.py", "f")) == "from the helper"


def test_same_named_scripts_are_two_modules_in_sys_modules(workbook, tmp_path):
    source = "import sys\n\ndef f():\n    return [__name__, sys.modules[__name__].__file__]\n"
    first = workbook(tmp_path / "uno" / "first.xlsx")
    second = workbook(tmp_path / "dos" / "second.xlsx")
    first.script("model.py", source)
    second.script("model.py", source)
    first.formula("A1", pycall("model.py", "f"))
    second.formula("A1", pycall("model.py", "f"))
    (one_name, one_file), (two_name, two_file) = first.values("A1:B1")[0], second.values("A1:B1")[0]
    assert one_name.startswith("pyinex_model_") and two_name.startswith("pyinex_model_") and one_name != two_name
    assert one_file.lower() == str(first.folder / "model.py").lower()
    assert two_file.lower() == str(second.folder / "model.py").lower()


def test_a_helper_imported_inside_a_function(book):
    (book.folder / "helper_late.py").write_text("VALUE = 'imported late'\n", encoding="utf-8")
    book.script("late.py", "def f():\n    import helper_late\n    return helper_late.VALUE\n")
    assert book.formula("A1", pycall("late.py", "f")) == "imported late"


def test_the_scripts_folder_comes_first_on_sys_path(book):
    # A file beside the script named like a standard module hides that module.
    (book.folder / "colorsys.py").write_text("WHOSE = 'the one beside the script'\n", encoding="utf-8")
    book.script("shadows.py", "def f():\n    import colorsys\n    return colorsys.WHOSE\n")
    assert book.formula("A1", pycall("shadows.py", "f")) == "the one beside the script"


def test_two_scripts_share_a_helper(book):
    (book.folder / "helper_shared.py").write_text("seen = []\n", encoding="utf-8")
    book.script("writes.py", "import helper_shared\n\ndef f(x):\n    helper_shared.seen.append(x)\n    return len(helper_shared.seen)\n")
    book.script("reads.py", "import helper_shared\n\ndef f():\n    return ', '.join(helper_shared.seen)\n")
    assert book.formula("A1", pycall("writes.py", "f", '"one"')) == 1.0
    assert book.formula("A2", pycall("writes.py", "f", '"two"')) == 2.0
    assert book.formula("A3", pycall("reads.py", "f")) == "one, two"


TARGETS = """import functools


class Half(float):
    def __new__(cls, x):
        return super().__new__(cls, x / 2)


class Doubler:
    def __call__(self, x):
        return x * 2


double = Doubler()
add_ten = functools.partial(lambda a, b: a + b, 10)
biggest = max
"""


def test_a_callable_of_any_kind_can_be_called(book):
    book.script("targets.py", TARGETS)
    assert book.formula("A1", pycall("targets.py", "Half", "4")) == 2.0  # A class.
    assert book.formula("A2", pycall("targets.py", "double", "4")) == 8.0
    assert book.formula("A3", pycall("targets.py", "add_ten", "4")) == 14.0
    assert book.formula("A4", pycall("targets.py", "biggest", "4", "9", "2")) == 9.0


def test_a_dotted_name_is_not_split(book, cleared):
    book.script("dotted.py", "import functools\n")
    assert book.formula("A1", pycall("dotted.py", "functools.reduce")) == E("#VALUE!")
    said = "has no function functools.reduce"
    window = cleared(said)
    assert window and said in window.text
