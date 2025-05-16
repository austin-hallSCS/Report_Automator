from tkinter import Tk
from tkinter import filedialog

class FilePath:
    def __init__(self, fileIsOut, fileExtension = None, filePath = None):
        self.fileIsOut = fileIsOut
        self.fileExtension = fileExtension
        self.filePath = self.gui_file_path()

    def __str__(self):
        return f'{self.filePath}'

    # Opens a file browser and returns the file path
    def gui_file_path(self):
        Tk().withdraw()
        match self.fileIsOut:
            case  False:
                filename = filedialog.askopenfilename()
            case True:
                filename = filedialog.asksaveasfilename(defaultextension=self.fileExtension)

        return filename
