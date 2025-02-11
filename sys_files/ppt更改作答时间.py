import tkinter as tk
from PIL import Image
from PIL import ImageTk
win=tk.Tk()

# class Gui(tk.Tk):
#     def __init__(self):
#         super().__init__()
#         self.title("Figure dynamic show v1.01")
#         # self.geometry("1000x800+400+100")
#         self.mainGui()
#         # self.mainloop()
#
#     def mainGui(self):
#         label3 = tk.Label(self, text="请输入本题的作答时间（分钟）:")
#         label3.pack()
#         sv3 = tk.StringVar()
#         # sv3.trace("w", input_title_test)
#         entry_input3 = tk.Entry(self, textvariable=sv3)
#         entry_input3.pack()
#
#         image = Image.open(r"图片题目汇总/20231252K8.jpg")
#         # image = Image.open(r"图片题目汇总/%s.jpg"%title)
#         photo = ImageTk.PhotoImage(image)
#         label = tk.Label(self, image=photo)
#         label.image = photo  # in case the image is recycled
#         label.pack()
#         # label4 = tk.Label(self, image=photo)
#         # label4.pack()
import tkinter
import tkinter as tk
from PIL import Image
from PIL import ImageTk


class Gui:
    def __init__(self):
        titles=[]
        self.gui = tkinter.Toplevel()  # create gui window
        self.gui.title("Image Display")  # set the title of gui
        self.gui.geometry("800x600")  # set the window size of gui

        label3 = tk.Label(self.gui, text="请输入本题的作答时间（分钟）:")
        label3.pack()
        sv3 = tk.StringVar()
        # sv3.trace("w", input_title_test)
        entry_input3 = tk.Entry(self.gui, textvariable=sv3)
        entry_input3.pack()

        button=tk.Button(self.gui,)

        load = Image.open(r"图片题目汇总/20231252K8.jpg")  # open image from path
        img = ImageTk.PhotoImage(load)  # read opened image

        label1 = tkinter.Label(self.gui, image=img)  # create a label to insert this image
        label1.pack()  # set the label in the main window

        self.gui.mainloop()  # start mainloop


main = Gui()

main = Gui()
# main.mainloop()
win.mainloop()
