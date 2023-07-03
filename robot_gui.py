from tkinter import *
from tkinter import filedialog, ttk
from tkinter.messagebox import showinfo
from Automacao_lancamento_cupom_fiscal_GUI import CouponSubmitAutomation as Automation


class AutomationGui:
    def __init__(self):
        self.root = Tk()
        self.file_path = StringVar()
        self.cpf_number = StringVar()
        self.password = StringVar()
        self.entity = StringVar()
        self.browser_name = StringVar()
        self.gui_aesthetics()
        self.file_search_field()
        self.cpf()
        self.user_password()
        self.entity_name()
        self.browser()
        self.initiate_button()
        self.help_button()
        self.root.mainloop()

    def gui_aesthetics(self):
        screen_width = self.root.winfo_width()
        screen_height = self.root.winfo_height()
        app_width = 650
        app_height = 300
        x = app_width - screen_width
        y = app_height - screen_height
        self.root.geometry(f'{app_width}x{app_height}+{x}+{y}')
        self.root.title('Automação de lançamento de cupom fiscal')
        self.root.iconphoto(False, PhotoImage(file='humanoid_for_exe.png'))
        self.root.resizable(width=False, height=False)
        self.root.call('tk', 'scaling', 1.3)
        self.root.config()

    def file_search_field(self):
        Label(self.root, text='Selecione arquivo Excel:').place(x=10, y=10)
        Entry(self.root,
              width=60,
              textvariable=self.file_path,
              state='readonly').place(x=200, y=12)
        Button(self.root,
               text='Procurar',
               command=self.file_search).place(x=580, y=8)

    def file_search(self):
        self.file_path.set(filedialog.askopenfilename(initialdir='/',
                                                      title='Selecione arquivo excel',
                                                      filetypes=(('Excel Macro-Enabled Workbook', '*.xlsm'),
                                                                 ('Excel Workbook', '*.xlsx'),
                                                                 ('Excel 97-2003 Workbook', '*.xls'))))

    def cpf(self):
        def on_cpf_entry(event):
            self.cpf_number.set(cpf_number.get())

        Label(self.root, text='Digite o seu CPF:').place(x=10, y=50)
        cpf_number = Entry(self.root, width=72)
        cpf_number.bind('<FocusOut>', on_cpf_entry)
        cpf_number.place(x=200, y=49)

    def user_password(self):
        def on_password_entry(event):
            self.password.set(password_entry.get())
        Label(self.root, text='Digite a senha:').place(x=10, y=90)
        password_entry = Entry(self.root, width=55, show='*')
        password_entry.bind('<FocusOut>', on_password_entry)
        password_entry.place(x=200, y=89)

        def show_password():
            if password_entry['show'] == '*':
                password_entry['show'] = ''
            else:
                password_entry['show'] = '*'

        Checkbutton(self.root, text='Mostrar senha', command=show_password).place(x=540, y=86)

    def entity_name(self):
        def on_enetity_entry(event):
            self.entity.set(entity_entry.get())
        Label(self.root,
              text='Digite o nome da entidade: (incluir "Associação")',
              wraplength=170,
              justify=LEFT).place(x=10, y=130)
        entity_entry = Entry(self.root, width=72)
        entity_entry.bind('<FocusOut>', on_enetity_entry)
        entity_entry.place(x=200, y=137)

    def browser(self):
        def on_browser_select(event):
            self.browser_name.set(browser_select.get())

        Label(self.root,
              text='Selecione o navegador de preferência:',
              wraplength=199,
              justify=LEFT).place(x=10, y=180)
        browser_list = ['Google Chrome', 'Firefox', 'Microsoft Edge']
        browser_select = ttk.Combobox(self.root,
                                      values=browser_list,
                                      width=69,
                                      state='readonly')
        browser_select.bind('<<ComboboxSelected>>', on_browser_select)
        browser_select.place(x=199, y=187)

    def initiate_button(self):
        Button(self.root,
               text='Iniciar o lançamento de notinhas',
               command=self.initiate_automation_class,
               width=20,
               height=3,
               wraplength=150).place(x=330, y=227)

    def initiate_automation_class(self):
        Automation(self.file_path.get(),
                   self.cpf_number.get(),
                   self.password.get(),
                   self.entity.get(),
                   self.browser_name.get())

    def help_button(self):
        Button(self.root,
               text='Ajuda',
               command=self.readme_open,
               width=20,
               height=3).place(x=490, y=227)

    def readme_open(self):
        message = 'Para mais detalhes, acesse o arquivo "Leia-me.txt" salvo na pasta do programa.'
        showinfo(title='Ajuda', message=message)


if __name__ == '__main__':
    AutomationGui()
