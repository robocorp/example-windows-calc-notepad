import platform
from pathlib import Path

from robocorp import windows
from robocorp.tasks import setup, task


desktop = windows.Desktop()


class LOCATORS:  # gather here all the locators for convenience
    # Common locators.
    CALCULATOR = 'name:"Calculator"'
    NOTEPAD = 'class:"Notepad" subname:"Notepad"'

    # Platform specific ones.
    class WINDOWS_11:
        CALC_PLUS = 'id:"plusButton"'
        CALC_NUMBER = 'id:"num{}Button"'
        CALC_DISPLAY = 'id:"CalculatorResults"'
        NOTE_EDITOR = 'name:"Text editor"'
        NOTE_SAVE = 'name:"Save as"'

    class WINDOWS_SERVER:
        CALC_PLUS = 'type:"Button" name:"Add"'
        CALC_NUMBER = 'type:"Button" name:"{}"'
        CALC_DISPLAY = 'type:"Text" name:"Result"'
        NOTE_EDITOR = 'name:"Text Editor"'
        NOTE_SAVE = 'name:"Save As"'


def get_win_version() -> str:
    """Windows only utility which returns the current Windows major version."""
    # Windows terminal `ver` command is bugged, until that's fixed, check by build
    #  number. (the same applies for `platform.version()`)
    WINDOWS_10_VERSION = "10"
    WINDOWS_11_BUILD = 22000
    version_parts = platform.version().split(".")
    major = version_parts[0]
    if major == WINDOWS_10_VERSION and int(version_parts[2]) >= WINDOWS_11_BUILD:
        major = "11"

    return major


# Based on the Windows version, operate with a specific set of locators.
IS_WINDOWS_11 = get_win_version() == "11"
WIN_LOCATORS = (
    LOCATORS.WINDOWS_11 if IS_WINDOWS_11 else LOCATORS.WINDOWS_SERVER
)


@setup
def open_close_apps(task):
    """Open Calculator and Notepad to automate them, then close both at the end."""
    desktop.windows_run("calc.exe")
    desktop.windows_run("notepad.exe")
    yield
    both_apps = f"{LOCATORS.CALCULATOR} or {LOCATORS.NOTEPAD}"
    desktop.close_windows(both_apps)


def add_numbers_with_calculator(*numbers: int) -> str:
    """Add an arbitrary list of numbers and return their result."""
    calc_window = windows.find_window(LOCATORS.CALCULATOR)
    plus_button = calc_window.find(WIN_LOCATORS.CALC_PLUS)
    for number in numbers:
        calc_window.click(WIN_LOCATORS.CALC_NUMBER.format(number))
        plus_button.click()
    display = calc_window.find(WIN_LOCATORS.CALC_DISPLAY)
    if IS_WINDOWS_11:
        # Pull the total from "Display is <nr>".
        result = display.name.rsplit(" ", 1)[-1]
    else:
        # Should support the "Value" pattern.
        result = str(display.get_value())
    return result


def save_text_with_notepad(text: str, *, path: Path):
    """Writes the result in the text editor then saves it in a text file."""
    notepad_window = windows.find_window(LOCATORS.NOTEPAD)
    text_edit = notepad_window.find(WIN_LOCATORS.NOTE_EDITOR)
    text_edit.set_value(text)
    save_as = "{Ctrl}{Shift}s" if IS_WINDOWS_11 else "{Ctrl}s"
    notepad_window.send_keys(save_as)
    save_as_window = notepad_window.find_child_window(WIN_LOCATORS.NOTE_SAVE)
    save_as_window.send_keys("{LAlt}n{Back}")
    save_as_window.send_keys(str(path.absolute()), send_enter=True)
    confirmation_window = save_as_window.find_child_window(
        "name:Confirm Save As", raise_error=False, timeout=1
    )
    if confirmation_window:
        # Overwrite existing file.
        confirmation_window.send_keys("{LAlt}y")


@task
def compute_numbers():
    """Do a computation in Calculator and store the result with Notepad."""
    result = add_numbers_with_calculator(1, 2, 3, 4)
    output_file = Path(".") / "output" / "result.txt"
    output_file.parent.mkdir(exist_ok=True)
    save_text_with_notepad(result, path=output_file)
