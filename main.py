from tkinter import Tk
from gui import PDFToWordGUI


def main():
    root = Tk()

    app = PDFToWordGUI(root)

    root.mainloop()


if __name__ == "__main__":
    main()