from robocorp import windows
from robocorp.tasks import task


desktop = windows.Desktop()


@task
def compute_numbers():
    """Do a computation in Calculator and store the result with Notepad."""
