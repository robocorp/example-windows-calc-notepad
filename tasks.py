from pathlib import Path

from robocorp import windows
from robocorp.tasks import setup, task


desktop = windows.Desktop()


class LOCATORS:
    CALCULATOR = "name:Calculator"
    NOTEPAD = "class:Notepad subname:Notepad"


@setup
def open_apps(task):
    """Open Calculator and Notepad to automate them, then close both at the end."""
    desktop.windows_run("calc.exe")
    desktop.windows_run("notepad.exe")
    yield
    both_apps = f"({LOCATORS.CALCULATOR}) or ({LOCATORS.NOTEPAD})"
    desktop.close_windows(both_apps)


def add_numbers_with_calculator(*numbers: int) -> str:
    """Add an arbitrary list of numbers and return their result."""
    calc_window = windows.find_window(LOCATORS.CALCULATOR)
    plus_button = calc_window.find("id:plusButton")
    for number in numbers:
        nr_button = calc_window.find(f"id:num{number}Button")
        nr_button.click()
        plus_button.click()
    display = calc_window.find("id:CalculatorResults")
    result = display.name.rsplit(" ", 1)[-1]  # pull the total from "Display is <nr>"
    return result


def save_text_with_notepad(text: str, *, path: Path):
    """Writes the result in the text editor then saves it in a text file."""
    notepad_window = windows.find_window(LOCATORS.NOTEPAD)
    text_edit = notepad_window.find("name:Text editor")
    text_edit.set_value(text)
    notepad_window.send_keys("{Ctrl}{Shift}s")
    save_as_window = notepad_window.find_child_window("name:Save as")
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
    output_file = Path(".") / "output" / "results.txt"
    output_file.parent.mkdir(exist_ok=True)
    save_text_with_notepad(result, path=output_file)
