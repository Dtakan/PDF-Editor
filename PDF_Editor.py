
import os
import PyPDF2
import os.path
from tkinter import *
from functools import partial
from tkinter import filedialog
from PyPDF2 import PdfFileReader
from tkinter import ttk, messagebox
from PyPDF2 import PdfFileWriter


class PDF_Editor:
    def __init__(self, root):
        self.window = root
        self.window.geometry("740x480")
        self.window.title('PDF Editor')

        # Color Options
        self.color_1 = "#472F17"
        self.color_2 = "#FFB266"
        self.color_3 = "black"
        self.color_4 = 'orange red'

        # Font Options
        self.font_1 = "Helvetica"
        self.font_2 = "Times New Roman"
        self.font_3 = "Kokila"

        self.saving_location = ''

        # ==============================================
        # ================Menubar Section===============
        # ==============================================
        # Creating Menubar
        self.menubar = Menu(self.window)

        # Adding Edit Menu and its sub menus
        edit = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='Edit', menu=edit)
        edit.add_command(label='PDF trennen', command=partial(self.SelectPDF, 1))
        edit.add_command(label='PDFs zusammenführen', command=self.Merge_PDFs_Data)
        edit.add_separator()
        edit.add_command(label='PDF Seiten drehen', command=partial(self.SelectPDF, 2))

        # Adding About Menu
        about = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='About', menu=about)
        about.add_command(label='About', command=self.AboutWindow)

        # Exit the Application
        exit = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='Exit', menu=exit)
        exit.add_command(label='Exit', command=self.Exit)

        # Configuring the menubar
        self.window.config(menu=self.menubar)
        # ===================End=======================

        # Creating a Frame
        self.frame_1 = Frame(self.window, bg= '#856ff8', width=740, height=480)
        self.frame_1.place(x=0, y=0)
        # Calling Home Page Window
        self.Home_Page()

    # ==============================================
    # =============Miscellaneous Functions==========
    # ==============================================
    def AboutWindow(self):
        messagebox.showinfo("PDF Editor", \
                            "PDF Editor\nDeveloped by Atakan Taşdirek (who's not very good at Graphic Design)")

    # Remove all widgets from the Home Page
    def ClearScreen(self):
        for widget in self.frame_1.winfo_children():
            widget.destroy()

    # It updates the current saving path label
    # with '/' (current working directory)
    def Update_Path_Label(self):
        self.path_label.config(text=self.saving_location)

    # After performing the rotation operation this
    # function gets called
    def Update_Rotate_Page(self):
        self.saving_location = ''
        self.ClearScreen()
        self.Home_Page()

    # It destroy the main GUI window of the
    # application
    def Exit(self):
        self.window.destroy()

    # ===================End========================

    # Home Page: It consists Three Buttons
    def Home_Page(self):
        self.ClearScreen()
        # ==============================================
        # ================Buttons Section===============
        # ==============================================
        # Split Button
        self.split_button = Button(self.frame_1, text='Seiten Trennen', \
                                   font=(self.font_1, 18, 'bold'), bg="#FFB266", fg="black", width=19, \
                                   command=partial(self.SelectPDF, 1))
        self.split_button.place(x=260, y=80)

        # Merge Button
        self.merge_button = Button(self.frame_1, text='PDFs zusammenführen', \
                                   font=(self.font_1, 18, 'bold'), bg="#FFB266", fg="black", \
                                   width=19, command=self.Merge_PDFs_Data)
        self.merge_button.place(x=260, y=160)

        # Rotation Button
        self.rotation_button = Button(self.frame_1, text='Seiten drehen', \
                                      font=(self.font_1, 18, 'bold'), bg="#FFB266", fg="black", \
                                      width=19, command=partial(self.SelectPDF, 2))
        self.rotation_button.place(x=260, y=240)

        # Pictures to PDF muss noch gemacht werden
        #self.convert_button = Button(self.frame_1, text='Bilder in PDF\nkonvertieren', \
         #                             font=(self.font_1, 19, 'bold'), bg="yellow", fg = "black", \
          #                            width=19, command=partial(self.ConvertPDF))
        #self.convert_button.place (x=260, y = 320)
        # ===================End=======================

    # Select the PDF for Splitting and Rotating
    def SelectPDF(self, to_call):
        self.PDF_path = filedialog.askopenfilename(initialdir="/", \
                                                   title="Wähle eine PDF Datei aus", filetypes=(("PDF files", "*.pdf*"),))
        if len(self.PDF_path) != 0:
            if to_call == 1:
                self.Split_PDF_Data()
            else:
                self.Rotate_PDFs_Data()

    # Select PDF files only for merging
    def SelectPDF_Merge(self):
        self.PDF_path = filedialog.askopenfilenames(initialdir="/", \
                                                    title="Wähle eine PDF Datei aus", filetypes=(("PDF files", "*.pdf*"),))
        for path in self.PDF_path:
            self.PDF_List.insert((self.PDF_path.index(path) + 1), path)

    # Select the directory where the result PDF
    # file/files will be stored
    def Select_Directory(self):
        # Storing the 'saving location' for the result file
        self.saving_location = filedialog.askdirectory(title= \
                                                           "Wähle ein Dateipfad")
        self.Update_Path_Label()

    # Get the data from the user for splitting a PDF file
    def Split_PDF_Data(self):
        pdfReader = PyPDF2.PdfFileReader(self.PDF_path)
        total_pages = pdfReader.numPages

        self.ClearScreen()
        # Button for getting back to the Home Page
        home_btn = Button(self.frame_1, text="Home", \
                          font=(self.font_1, 10, 'bold'), command=self.Home_Page)
        home_btn.place(x=10, y=10)

        # Header Label
        header = Label(self.frame_1, text="PDF Seiten teilen", \
                       font=(self.font_3, 25, "bold"), bg=self.color_2, fg=self.color_1)
        header.place(x=265, y=15)

        # Label for showing the total number of pages
        self.pages_label = Label(self.frame_1, \
                                 text=f"Seitenanzahl: {total_pages}", \
                                 font=(self.font_2, 20, 'bold'), bg=self.color_2, fg=self.color_3)
        self.pages_label.place(x=40, y=70)

        # From Label: the page number from where the
        # user want to split the PDF pages
        From = Label(self.frame_1, text="Von", \
                     font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        From.place(x=40, y=120)

        self.From_Entry = Entry(self.frame_1, font=(self.font_2, 12, 'bold'), \
                                width=8)
        self.From_Entry.place(x=40, y=160)

        # To Label
        To = Label(self.frame_1, text="Bis", font=(self.font_2, 16, 'bold'), \
                   bg=self.color_2, fg=self.color_1)
        To.place(x=160, y=120)

        self.To_Entry = Entry(self.frame_1, font=(self.font_2, 12, 'bold'), \
                              width=8)
        self.To_Entry.place(x=160, y=160)

        Cur_Directory = Label(self.frame_1, text="Speicherort", \
                              font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        Cur_Directory.place(x=300, y=120)

        # Constant
        self.path_label = Label(self.frame_1, text='/', \
                                font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_3)
        self.path_label.place(x=300, y=160)

        # Button for selecting the directory
        # where the splitted PDFs will be stored
        select_loc_btn = Button(self.frame_1, text="Wähle ein Dateipfad", \
                                font=(self.font_1, 8, 'bold'), command=self.Select_Directory)
        select_loc_btn.place(x=320, y=200)

        split_button = Button(self.frame_1, text="Teilen", \
                              font=(self.font_3, 16, 'bold'), bg=self.color_4, fg=self.color_1, \
                              width=12, command=self.Split_PDF)
        split_button.place(x=250, y=250)

    # Get the data from the user for Merge PDF files
    def Merge_PDFs_Data(self):
        self.ClearScreen()
        # Button for get back to the Home Page
        home_btn = Button(self.frame_1, text="Home", \
                          font=(self.font_1, 10, 'bold'), command=self.Home_Page)
        home_btn.place(x=10, y=10)

        # Header Label
        header = Label(self.frame_1, text="PDFs zusammenführen", \
                       font=(self.font_3, 25, "bold"), bg=self.color_2, fg=self.color_1)
        header.place(x=265, y=15)

        select_pdf_label = Label(self.frame_1, text="Wähle PDFs aus", \
                                 font=(self.font_2, 20, 'bold'), bg=self.color_2, fg=self.color_3)
        select_pdf_label.place(x=40, y=70)

        open_button = Button(self.frame_1, text="Ordner öffnen", \
                             font=(self.font_1, 9, 'bold'), command=self.SelectPDF_Merge)
        open_button.place(x=55, y=110)

        Cur_Directory = Label(self.frame_1, text="Speicherort", \
                              font=(self.font_2, 19, 'bold'), bg=self.color_2, fg=self.color_1)
        Cur_Directory.place(x=40, y=150)

        # Constant
        self.path_label = Label(self.frame_1, text='/', \
                                font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_3)
        self.path_label.place(x=40, y=190)

        # Button for selecting the directory
        # where the merged PDFs will be stored
        select_loc_btn = Button(self.frame_1, text="Wähle ein Dateipfad", \
                                font=(self.font_1, 9, 'bold'), command=self.Select_Directory)
        select_loc_btn.place(x=55, y=225)

        saving_name = Label(self.frame_1, text="Benenne die PDF", \
                            font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        saving_name.place(x=40, y=270)

        # Get the 'result file' name from the user
        self.sv_name_entry = Entry(self.frame_1, \
                                   font=(self.font_2, 12, 'bold'), width=20)
        self.sv_name_entry.insert(0, 'Result')
        self.sv_name_entry.place(x=40, y=310)

        merge_btn = Button(self.frame_1, text="Zusammenführen", \
                           font=(self.font_1, 10, 'bold'), command=self.Merge_PDFs)
        merge_btn.place(x=80, y=350)

        listbox_label = Label(self.frame_1, text="Ausgewählte PDFs", \
                              font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        listbox_label.place(x=400, y=72)

        # Listbox for showing the selected PDF files
        self.PDF_List = Listbox(self.frame_1, width=40, height=15)
        self.PDF_List.place(x=400, y=110)

        delete_button = Button(self.frame_1, text="Löschen", \
                               font=(self.font_1, 9, 'bold'), command=self.Delete_from_ListBox)
        delete_button.place(x=400, y=395)

        more_button = Button(self.frame_1, text="Mehr auswählen", \
                             font=(self.font_1, 9, 'bold'), command=self.SelectPDF_Merge)
        more_button.place(x=480, y=395)

    # Get the data from the user for Rotating one/multiple
    # pages of a PDF file
    def Rotate_PDFs_Data(self):
        self.ClearScreen()

        pdfReader = PyPDF2.PdfFileReader(self.PDF_path)
        total_pages = pdfReader.numPages

        # Button for get back to the Home Page
        home_btn = Button(self.frame_1, text="Home", \
                          font=(self.font_1, 10, 'bold'), command=self.Home_Page)
        home_btn.place(x=10, y=10)

        # Header Label
        header = Label(self.frame_1, text="PDFs rotieren", \
                       font=(self.font_3, 25, "bold"), bg=self.color_2, fg=self.color_1)
        header.place(x=265, y=15)

        # Label for showing the total number of pages
        self.pages_label = Label(self.frame_1, \
                                 text=f"Seitenanzahl: {total_pages}", \
                                 font=(self.font_2, 20, 'bold'), bg=self.color_2, fg=self.color_3)
        self.pages_label.place(x=40, y=90)

        Cur_Directory = Label(self.frame_1, text="Speicherort", \
                              font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        Cur_Directory.place(x=40, y=150)

        self.fix_label = Label(self.frame_1, \
                               text="Rotiere diese Seiten (z.B. 2, 4 für Seite 2 & 4)", \
                               font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        self.fix_label.place(x=260, y=150)

        self.fix_entry = Entry(self.frame_1, \
                               font=(self.font_2, 12, 'bold'), width=40)
        self.fix_entry.place(x=260, y=190)

        # Constant
        self.path_label = Label(self.frame_1, text='/', \
                                font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_3)
        self.path_label.place(x=40, y=190)

        # Button for selecting the directory
        # where the rotated PDFs will be stored
        select_loc_btn = Button(self.frame_1, text="Speicherort auswählen", \
                                font=(self.font_1, 9, 'bold'), command=self.Select_Directory)
        select_loc_btn.place(x=55, y=225)

        saving_name = Label(self.frame_1, text="Benenne die PDF", \
                            font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        saving_name.place(x=40, y=270)

        # Get the 'result file' name from the user
        self.sv_name_entry = Entry(self.frame_1, \
                                   font=(self.font_2, 12, 'bold'), width=20)
        self.sv_name_entry.insert(0, 'Result')
        self.sv_name_entry.place(x=40, y=310)

        which_side = Label(self.frame_1, text="Seitenausrichtung", \
                           font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        which_side.place(x=260, y=230)

        # Rotation Alignment(Clockwise and Anti-Clockwise)
        text = StringVar()
        self.alignment = ttk.Combobox(self.frame_1, textvariable=text)
        self.alignment['values'] = ('im Uhrzeigersinn',
                                    'gegen den Uhrzeigersinn'
                                    )
        self.alignment.place(x=260, y=270)

        rotate_button = Button(self.frame_1, text="Rotieren", \
                               font=(self.font_3, 16, 'bold'), bg=self.color_4, \
                               fg=self.color_1, width=12, command=self.Rotate_PDFs)
        rotate_button.place(x=255, y=360)

    # It manages the task for Splitting the
    # selected PDF file
    def Split_PDF(self):
        if self.From_Entry.get() == "" and self.To_Entry.get() == "":
            messagebox.showwarning("Achtung!", \
                                   "Gebe die Seite(n) an,\n die du teilen möchtest")
        else:
            from_page = int(self.From_Entry.get()) - 1
            to_page = int(self.To_Entry.get())

            pdfReader = PyPDF2.PdfFileReader(self.PDF_path)

            for page in range(from_page, to_page):
                pdfWriter = PdfFileWriter()
                pdfWriter.addPage(pdfReader.getPage(page))

                splitPage = os.path.join(self.saving_location, f'{page + 1}.pdf')
                resultPdf = open(splitPage, 'wb')
                pdfWriter.write(resultPdf)

            resultPdf.close()
            messagebox.showinfo("Yaaaay", "Die PDF Datei wurde erfolgreich geteilt")

            self.saving_location = ''
            self.total_pages = 0
            self.ClearScreen()
            self.Home_Page()

    # It manages the task for Merging the
    # selected PDF files
    def Merge_PDFs(self):
        if len(self.PDF_path) == 0:
            messagebox.showerror("Fehler!", "Wähle eine PDF aus")
        else:
            if self.saving_location == '':
                curDirectory = os.getcwd()
            else:
                curDirectory = str(self.saving_location)

            presentFiles = list()

            for file in os.listdir(curDirectory):
                presentFiles.append(file)

            checkFile = f'{self.sv_name_entry.get()}.pdf'

            if checkFile in presentFiles:
                messagebox.showwarning('Achtung!', \
                                       "Wähle einen anderen Dateinamen aus")
            else:
                pdfWriter = PyPDF2.PdfFileWriter()

                for file in self.PDF_path:
                    pdfReader = PyPDF2.PdfFileReader(file)
                    numPages = pdfReader.numPages
                    for page in range(numPages):
                        pdfWriter.addPage(pdfReader.getPage(page))

                mergePage = os.path.join(self.saving_location, \
                                         f'{self.sv_name_entry.get()}.pdf')
                mergePdf = open(mergePage, 'wb')
                pdfWriter.write(mergePdf)

                mergePdf.close()
                messagebox.showinfo("Yaaaaay",
                                    "Die PDFs wurden erfolgreich zusammengefügt!")

                self.saving_location = ''
                self.ClearScreen()
                self.Home_Page()

    # Delete an item(One PDF Path)
    def Delete_from_ListBox(self):
        try:
            if len(self.PDF_path) < 1:
                messagebox.showwarning('Achtung!', \
                                       'Es gibt keine Dateien zum löschen')
            else:
                for item in self.PDF_List.curselection():
                    self.PDF_List.delete(item)

                self.PDF_path = list(self.PDF_path)
                del self.PDF_path[item]
        except Exception:
            messagebox.showwarning('Achtung!', "Wähle zuerst PDFs aus")

    # It manages the task for Rotating the pages/page of
    # the selected PDF file
    def Rotate_PDFs(self):
        need_to_fix = list()

        if self.fix_entry.get() == "":
            messagebox.showwarning("Achtung!", \
                                   "Trenne die Seitenzahl(en) mit einem Komma")
        else:
            for page in self.fix_entry.get().split(','):
                need_to_fix.append(int(page))

            if self.saving_location == '':
                curDirectory = os.getcwd()
            else:
                curDirectory = str(self.saving_location)

            presentFiles = list()

            for file in os.listdir(curDirectory):
                presentFiles.append(file)

            checkFile = f'{self.sv_name_entry.get()}.pdf'

            if checkFile in presentFiles:
                messagebox.showwarning('Achtung!', \
                                       "Wähle einen anderen Dateinamen aus")
            else:
                if self.alignment.get() == 'im Uhrzeigersinn':
                    pdfReader = PdfFileReader(self.PDF_path)
                    pdfWriter = PdfFileWriter()

                    rotatefile = os.path.join(self.saving_location, \
                                              f'{self.sv_name_entry.get()}.pdf')
                    fixed_file = open(rotatefile, 'wb')

                    for page in range(pdfReader.getNumPages()):
                        thePage = pdfReader.getPage(page)
                        if (page + 1) in need_to_fix:
                            thePage.rotateClockwise(90)

                        pdfWriter.addPage(thePage)

                    pdfWriter.write(fixed_file)
                    fixed_file.close()
                    messagebox.showinfo('Yippie', 'Rotation Complete')
                    self.Update_Rotate_Page()

                elif self.alignment.get() == 'gegen den Uhrzeigersinn':
                    pdfReader = PdfFileReader(self.PDF_path)
                    pdfWriter = PdfFileWriter()

                    rotatefile = os.path.join(self.saving_location, \
                                              f'{self.sv_name_entry.get()}.pdf')
                    fixed_file = open(rotatefile, 'wb')

                    for page in range(pdfReader.getNumPages()):
                        thePage = pdfReader.getPage(page)
                        if (page + 1) in need_to_fix:
                            thePage.rotateCounterClockwise(90)

                        pdfWriter.addPage(thePage)

                    pdfWriter.write(fixed_file)
                    fixed_file.close()
                    messagebox.showinfo('Hurra', 'Rotation Completed')
                    self.Update_Rotate_Page()
                else:
                    messagebox.showwarning('Achtung!', \
                                           "Wähle eine Ausrichtung")


# The main function
if __name__ == "__main__":
    root = Tk()
    # Creating a CountDown class object
    obj = PDF_Editor(root)
    root.mainloop()