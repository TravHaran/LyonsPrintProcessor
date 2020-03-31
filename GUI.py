import tkinter as tk
import imaplib
import gspread
from gspread_formatting import *
from httplib2 import ServerNotFoundError
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.multipart import MIMEMultipart
import smtplib
from datetime import date


# install:oauth2client, gspread, PyOpenSSL, gspread-formatting
# noinspection PyAttributeOutsideInit,SpellCheckingInspection,DuplicatedCode
class Window:
    def __init__(self):
        self.scope = ['https://spreadsheets.google.com/feeds',
                      'https://www.googleapis.com/auth/drive']
        self.credsdict = {'type': 'service_account',
                          'project_id': 'lyons-email-updatesheet-script',
                          'private_key_id': '2414046f0774321dd439b121749c6eff1db8713b',
                          'private_key': '-----BEGIN PRIVATE '
                                         'KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDrHJwzqsP88DVg'
                                         '\nHplpDCpKYwUYGUEbRg0sqFx7rJy9mYOXD7/TpducJlInQs5PbTiF/yQ0o/IjAaSB'
                                         '\nWWqaQbCv+COSVg5udAcbrP6xSTv+3L2crqFMDXnHbKHHfhvNTQSyfh2JSIZY78le'
                                         '\nn6dcq0g3MvHEt7Qy89bdgC966XsXXVe+va0t8lI26IgPf9ZTSFm4a4TYStsT3mP3'
                                         '\nKl4RegTn03bMHYdPmOm2B/I5E814RG8Wjsk0FIYhW6reYNhLuBLrS14744bWmd6f'
                                         '\nYvZLpKmEf5EN91nD/veWoKB2qLhESfmpP1iipieoVpm+VvfDlWmcusl8s+w0/ETf'
                                         '\nkk1BMo5dAgMBAAECggEAE3cw744p3907bhPae7oIHlSIbXBZ1Zo9KP9feNXXvFLj'
                                         '\ndDRXm3xV7F2324xKbIUMcvum0bzpJUDTj+oJS3A44rjWqRz64OY2WHJAPAlmMDmy'
                                         '\ncTB8JkHPXVV/J3cnch34T5bldyJMDTz9HRp2ztNXjUpoffL/tmA93+TnCXQfPtXQ'
                                         '\n9vo1nj2n3Pix4zSZSwMk18Ll6vERbjKHffaSc/xdvMuHgEFa48ze3cOmrpFyJHM1'
                                         '\nrdZ6qhCUQPwgaHHFTgKcIDU5ILVlmuIm1DnUCo3K0ocDdov7zwyU/J/5gUoFzOpW'
                                         '\nCA7pwy2QbLZ3XWGL01gV6d1mGqiF9ajNxLBezu7AwQKBgQD4/44ywpno0nbEXlNj'
                                         '\n8zLpWGJyFkVDYwdPtcPlQZzPdq/UJd0bfApSnRaNDTdNwTbYtUSyT3QlXUOgzjai'
                                         '\nn2zKf6cXRtaGScGnmF4acJFvBSw9xTaaDaLs7IVoGD/CaCPwwJJ71OBLRTPxvIev'
                                         '\ntMfWWsANpYy/6pipuJI3Ujb/ZQKBgQDxuRdLN8Ox1OVQxAXUbykqyeuNt5h+h/vk'
                                         '\n9SdrYu72KmdHIUHgfu1L9N9yFp/KpGZ4CnhlkGX0LXoDnqQKd1yjiLxGrWgLk76P\n'
                                         '/dXCpLs4N7av94je5Z2e2Q1v+1qQ8cAs13UCcwemHe1WxRfJqsj8kP3ZX2FFGUke'
                                         '\nJgzZGpkPmQKBgDISWgsVHRQ3tpB4k3ZnApbwIiPlHJqXgHHkEHe6wQjrSiJ0Vslf'
                                         '\nIUhJtK46uSNWtmvPz/e3iJi275GXxl7fhmYWU4iXwy4QCPRl7I6OkoBr3uCxFvDV'
                                         '\nyyyvx4gOUEwM2yVf5FUoks4wJWj4S6TmysTtTO+xmeNCDt8acbTUQKENAoGAPH0k'
                                         '\n5x29SvMLr3peOxrWIm8FEyGud3tv/YubobPQOKnDznj0E0mv+CH/CH3A3uTk/4Uf'
                                         '\nO8s2uDPpJJ6+TiAwfnvpIYajUsJWHZJXu62dbCQFA2PeTGkJWIbYZf1wXHUishX4'
                                         '\nofRHJbq3ec84dK7YPNvLqmnD3ZbGRVUgQfP1+YECgYBiwykzHoHeY4e1yAbDzPzU'
                                         '\nwvlxE6bGgJXJLKObhRJjoFn5zyAEteMFOdTh7OUfcWIxi3/HdxIR8k8FvupNAOAi'
                                         '\nBQYLiMjIRLCJ3ljR3JHN3fDrwVpNOBEdtT1S7i8jWTA0tb2XSn9mN3IDjrdE5aqU'
                                         '\n2hm0MtKP0tWBhzfllvvOHg==\n-----END PRIVATE KEY-----\n',
                          'client_email': 'id-d-print-request@lyons-email-updatesheet-script.iam.gserviceaccount.com',
                          'client_id': '117542574436878496508',
                          'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
                          'token_uri': 'https://oauth2.googleapis.com/token',
                          'auth_provider_x509_cert_url': 'https://www.googleapis.com/oauth2/v1/certs',
                          'client_x509_cert_url': 'https://www.googleapis.com/robot/v1/metadata/x509/id-d-print'
                                                  '-request%40lyons-email-updatesheet-script.iam.gserviceaccount.com'}

        self.fmtreadypickup = CellFormat(
            backgroundColor=Color(0.078, 0.616, 1),
        )
        self.fmtdenied = CellFormat(
            backgroundColor=Color(0.878, 0.4, 0.4),
        )
        self.fmtfailed = CellFormat(
            backgroundColor=Color(0.957, 0.8, 0.8),
        )
        self.fmtclarification = CellFormat(
            backgroundColor=Color(1, 0.851, 0.4),
        )
        self.fmtcancelled = CellFormat(
            backgroundColor=Color(0.71, 0.37, 0.02)
        )
        self.fmtpickedup = CellFormat(
            backgroundColor=Color(0.85, 0.85, 0.85),
        )
        self.fmtneverpickedup = CellFormat(
            backgroundColor=Color(0.4, 0.31, 0.65),
        )
        # define message for email
        self.msg = MIMEMultipart()
        # Setup window
        self.window = tk.Tk()
        self.window.geometry("400x150+800+300")
        self.window.resizable(0, 0)
        #self.window.iconbitmap(r"favicon.ico")
        self.titleFrame = ""

        # creating a menu instance
        menu = tk.Menu(self.window)
        self.window.config(menu=menu)
        file = tk.Menu(menu)
        file.add_command(label="About", command=self.about)
        menu.add_cascade(label="File", menu=file)

        # Initiate Login
        self.LoginMenu()
        self.window.mainloop()

    def about(self):
        self.aboutWindow = tk.Tk()
        self.aboutWindow.geometry("300x75")
        self.aboutWindow.title("3D Print Request - About")
        self.aboutWindow.resizable(0, 0)
        self.aboutFrame = tk.Frame(self.aboutWindow)
        self.aboutFrame.pack()
        tk.Label(self.aboutFrame, text="3D Print Request V.1").grid(row=0, column=0)
        tk.Label(self.aboutFrame, text="By: Travis Ratnaharan & Ridvan Song").grid(row=1, column=0)
        tk.Label(self.aboutFrame, text="Built with Python").grid(row=2, column=0)

    def LoginMenu(self):
        password = tk.StringVar()
        self.window.title("3D Print Request - Login")
        self.titleFrame = tk.Frame(self.window)
        self.titleFrame.pack()
        Account = 'lyons.newmedia@gmail.com'
        Sender = tk.StringVar(self.titleFrame, value=Account)
        tk.Label(self.titleFrame, text="Account:").grid(row = 0, column = 0, pady = 10, padx = 20, sticky = "E")
        tk.Entry(self.titleFrame, textvariable=Sender, width=30).grid(row = 0, column = 1, pady = 10, padx = 20)
        tk.Label(self.titleFrame, text='Enter Password:').grid(row = 1, column = 0, pady = 10, padx = 20, sticky = "E")
        tk.Entry(self.titleFrame, textvariable=password, show='•', width=30).grid(row = 1, column = 1, pady = 10, padx = 20)
        bt = tk.Button(self.titleFrame, text='Enter', command=lambda: self.PasswordEntry(Sender, password), padx = 30)
        bt.grid(row = 2, column = 0, columnspan = 2)
        self.window.bind('<Return>', lambda event=None: bt.invoke())

    def PasswordEntry(self, Sender, password):
        self.User = str(Sender.get())
        try:
            self.Password = str(password.get())
        except UnicodeDecodeError:
            statusLabel = tk.Label(self.titleFrame, text="Login Failed, Invalid Email/Password", fg="red")
            statusLabel.grid(row=3, column=0, columnspan=2)
            statusLabel.update()
            self.titleFrame.after(1000, statusLabel.destroy())
            self.titleFrame.destroy()
            # Re-Initiate Login
            self.LoginMenu()
        self.Authorize()
        if self.wifi == 1:
            # Exception Handling for Invalid Login Credentials
            try:
                self.mail.login(self.User, self.Password)
            except Exception as e:
                self.log = '0'
            else:
                self.log = '1'
            if self.log == '0':
                statusLabel = tk.Label(self.titleFrame, text="Login Failed, Invalid Email/Password",fg="red")
                statusLabel.grid(row = 3, column = 0, columnspan = 2)
                statusLabel.update()
                self.titleFrame.after(1000, statusLabel.destroy())
                self.titleFrame.destroy()
                # Re-Initiate Login
                self.LoginMenu()
            elif self.log == '1':
                statusLabel = tk.Label(self.titleFrame, text="Login Success",fg="green")
                statusLabel.grid(row = 3, column = 0, columnspan = 2)
                statusLabel.update()
                self.titleFrame.after(1000, statusLabel.destroy())
                self.titleFrame.destroy()
                # Initiate Menu
                self.backToMenu()

    def Authorize(self):
        self.credentials = ServiceAccountCredentials.from_json_keyfile_dict(self.credsdict, self.scope)
        # Exception Handling for no Internet Connection
        try:
            self.gc = gspread.authorize(self.credentials)
        except ServerNotFoundError:
            connection = tk.Label(self.titleFrame, text="Login Failed, No Internet Connection",fg="red")
            connection.pack()
            connection.update()
            self.titleFrame.after(1000, connection.destroy())
            self.titleFrame.destroy()
            # Re-Initiate Login
            self.LoginMenu()
            self.wifi = 0
        else:
            self.sh = self.gc.open('3D Printing Requests')
            self.worksheet_list = self.sh.worksheets()
            self.worksheet_str = [str(i) for i in self.worksheet_list]
            self.worksheet = [x.replace("<Worksheet '", '').split("' id:", 1)[0] for x in self.worksheet_str]
            self.mail = imaplib.IMAP4_SSL('imap.gmail.com')
            self.wifi = 1

    def StartMenu(self):
        self.window.title("3D Print Request - Menu")
        self.window.geometry("500x500")
        self.titleFrame = tk.Frame(self.window)
        self.titleFrame.pack()
        tk.Label(self.titleFrame, text="Select the time period of the Print Request").pack()
        self.workSheet = tk.StringVar(self.titleFrame)
        self.workSheet.set(list(self.worksheet)[0])  # default value
        tk.OptionMenu(self.titleFrame, self.workSheet, *self.worksheet).pack()
        tk.Label(self.titleFrame, text="").pack()
        tk.Button(self.titleFrame, text="New Submission Processing", width="40", pady="5",
                  command=lambda: self.getInfo(self.defineNewPatronInfo, "Submit", "3D Print Request - New Submission",
                                               0)).pack()
        tk.Button(self.titleFrame, text="Ready For Pickup", width="40", pady="5",
                  command=lambda: self.getInfo(self.readyForPickup, "Send Email",
                                               "3D Print Request - Ready For Pickup", 1)).pack()
        tk.Button(self.titleFrame, text="Delayed Printing", width="40", pady="5",
                  command=lambda: self.getInfo(self.DelayedPrinting, "Send Email",
                                               "3D Print Request - Delayed Printing", 4)).pack()
        tk.Button(self.titleFrame, text="Denied", width="40", pady="5",
                  command=lambda: self.getInfo(self.Denied, "Send Email", "3D Print Request - Denied", 2)).pack()
        tk.Button(self.titleFrame, text="Dimensions Clarification - Skewed Print", width="40", pady="5",
                  command=lambda: self.getInfo(self.Clarification_Skewed, "Send Email",
                                               "3D Print Request - Skewed Print", 1)).pack()
        tk.Button(self.titleFrame, text="Dimensions Clarification - Large Print", width="40", pady="5",
                  command=lambda: self.getInfo(self.Clarification_Large, "Send Email",
                                               "3D Print Request - Large Print", 1)).pack()
        tk.Button(self.titleFrame, text="Reminder", width="40", pady="5",
                  command=lambda: self.getInfo(self.Reminder, "Send Email", "3D Print Request - Reminder", 3)).pack()
        tk.Button(self.titleFrame, text="Failed", width="40", pady="5",
                  command=lambda: self.getInfo(self.Failed, "Send Email", "3D Print Request - Failed", 2)).pack()
        tk.Button(self.titleFrame, text="Picked Up", width="40", pady="5",
                  command=lambda: self.getInfo(self.pickedUp, "Update Spreadsheet",
                                               "3D Print Request - Picked Up", 1)).pack()
        tk.Button(self.titleFrame, text="Never Picked Up", width="40", pady="5",
                  command=lambda: self.getInfo(self.nevPickedUp, "Update Spreadsheet",
                                               "3D Print Request - Never Picked Up", 1)).pack()
        tk.Button(self.titleFrame, text="Cancelled", width="40", pady="5",
                  command=lambda: self.getInfo(self.cancelled, "Update Spreadsheet",
                                               "3D Print Request - Cancelled", 2)).pack()

    def backToMenu(self):
        self.wks = ""
        self.row_number = ""
        self.rowstr = ""
        self.z = ""


        # Reset User Input Fields
        self.name = ""
        self.Ticketnum = ""
        self.patron_email = ""
        self.StaffInitials = ""
        self.dateToday = ""
        self.CourseYN = ""
        self.CourseCode = ""
        self.affiliation = ""
        self.department = ""
        self.research = ""
        self.OwnC = ""
        self.consent = ""
        self.handle = ""
        self.SD = ""
        self.Fname = ""
        self.Ptime = ""
        self.reasonEntry = ""
        self.dateEntry1 = ""
        self.dateEntry2 = ""
        self.responseDate = ""

        self.StartMenu()

    def getInfo(self, function, text, title, option):
        self.window.geometry("500x225")
        self.window.title(title)
        self.gc.login()
        self.wks = self.sh.get_worksheet(list(self.worksheet).index(self.workSheet.get()))
        self.titleFrame.destroy()
        self.infoFrame = tk.Frame(self.window)
        self.infoFrame.pack()
        self.nameEntry = tk.StringVar(self.infoFrame, value=self.name)
        self.ticketNumEntry = tk.StringVar(self.infoFrame, value=self.Ticketnum)
        self.emailEntry = tk.StringVar(self.infoFrame, value=self.patron_email)
        self.StaffInitials = tk.StringVar(self.infoFrame)
        self.dateToday = tk.StringVar(self.infoFrame, value=date.today().strftime("%m/%d/%Y"))
        self.CourseYN = tk.IntVar(self.infoFrame)
        self.CourseCode = tk.StringVar(self.infoFrame)
        self.affiliation = tk.StringVar(self.infoFrame)
        self.department = tk.StringVar(self.infoFrame)
        self.research = tk.IntVar(self.infoFrame)
        self.OwnC = tk.IntVar(self.infoFrame)
        self.consent = tk.IntVar(self.infoFrame)
        self.handle = tk.StringVar(self.infoFrame)
        self.SD = tk.StringVar(self.infoFrame)
        self.Fname = tk.StringVar(self.infoFrame)
        self.Ptime = tk.StringVar(self.infoFrame)
        self.reasonEntry = tk.StringVar()
        userDateString = 'mm/dd/yyyy'
        self.dateEntry1 = tk.StringVar(self.infoFrame, value=userDateString)
        self.dateEntry2 = tk.StringVar(self.infoFrame, value=userDateString)
        userDateString2 = 'Month, DD, YYYY'
        self.responseDate = tk.StringVar(self.infoFrame, value=userDateString2)
        tk.Label(self.infoFrame, text=title,font="10",padx=80,pady=10).grid(row=0,column = 0,columnspan=3)
        tk.Label(self.infoFrame, text="Enter Ticket #:",pady=10).grid(row = 1,column = 0,sticky='e')
        tk.Entry(self.infoFrame, textvariable=self.ticketNumEntry).grid(row=1,column = 1,sticky='w')
        self.ticket = str(self.ticketNumEntry)
        search = tk.Button(self.infoFrame, text="Search",
                  command=lambda: self.findTicket(function, text, title, option))
        search.grid(row=1,column = 2,sticky='w')
        self.window.bind('<Return>', lambda event=None: search.invoke())

        tk.Label(self.infoFrame, text="Enter Patron Name:",pady=10).grid(row=2,column=0,sticky="E")
        tk.Entry(self.infoFrame, textvariable=self.nameEntry).grid(row=2,column=1,sticky='w')
        tk.Label(self.infoFrame, text="Enter Patron Email:",pady=10).grid(row=3,column=0,sticky="E")
        tk.Entry(self.infoFrame, textvariable=self.emailEntry,width = 30).grid(row=3,column=1,sticky='w')
        if option == 0:
            self.window.geometry("500x750")
            tk.Label(self.infoFrame, text="Date:",pady=10).grid(row=4,column=0,sticky="E")
            tk.Entry(self.infoFrame, textvariable=self.dateToday).grid(row=4,column=1,sticky='w')
            tk.Label(self.infoFrame, text="Staff Initials:",pady=10).grid(row=5,column=0,sticky="E")
            tk.Entry(self.infoFrame, textvariable=self.StaffInitials).grid(row=5,column=1,sticky='w')
            tk.Label(self.infoFrame, text="Is it for a course?",pady=10).grid(row=6,column=0,sticky="E")
            CourseYNF = tk.LabelFrame(self.infoFrame)
            CourseYNF.grid(row=6,column=1,sticky='w')
            tk.Radiobutton(CourseYNF, text="Yes", padx=20, variable=self.CourseYN, value=1).pack(side="left")
            tk.Radiobutton(CourseYNF, text="No", padx=20, variable=self.CourseYN, value=0).pack(side="left")
            tk.Label(self.infoFrame, text="Course Code:",pady=10).grid(row=7,column=0,sticky="E")
            tk.Entry(self.infoFrame, textvariable=self.CourseCode).grid(row=7,column=1,sticky='w')
            tk.Label(self.infoFrame, text="Affiliation:",pady=10).grid(row=8,column=0,sticky="E")
            tk.Entry(self.infoFrame, textvariable=self.affiliation).grid(row=8,column=1,sticky='w')
            tk.Label(self.infoFrame, text="Department:",pady=10).grid(row=9,column=0,sticky="E")
            tk.Entry(self.infoFrame, textvariable=self.department).grid(row=9,column=1,sticky='w')
            tk.Label(self.infoFrame, text="Is it for research?",pady=10).grid(row=10,column=0,sticky="E")
            researchF = tk.LabelFrame(self.infoFrame)
            researchF.grid(row=10,column=1,sticky='w')
            tk.Radiobutton(researchF, text="Yes", padx=20, variable=self.research, value=1).pack(side="left")
            tk.Radiobutton(researchF, text="No", padx=20, variable=self.research, value=0).pack(side="left")
            tk.Label(self.infoFrame, text="Did you create this model?",pady=10).grid(row=11,column=0,sticky="E")
            createF = tk.LabelFrame(self.infoFrame)
            createF.grid(row=11,column=1,sticky='w')
            tk.Radiobutton(createF, text="Yes", padx=20, variable=self.OwnC, value=1).pack(side="left")
            tk.Radiobutton(createF, text="No", padx=20, variable=self.OwnC, value=0).pack(side="left")
            tk.Label(self.infoFrame, text="Do you consent to \nInstagram Post?",pady=10).grid(row=12,column=0,sticky="E")
            consentF = tk.LabelFrame(self.infoFrame)
            consentF.grid(row=12,column=1,sticky='w')
            tk.Radiobutton(consentF, text="Yes", padx=20, variable=self.consent, value=1).pack(side="left")
            tk.Radiobutton(consentF, text="No", padx=20, variable=self.consent, value=0).pack(side="left")
            tk.Label(self.infoFrame, text="Instagram Handle:",pady=10).grid(row=13,column=0,sticky="E")
            tk.Entry(self.infoFrame, textvariable=self.handle).grid(row=13,column=1,sticky='w')
            tk.Label(self.infoFrame, text="File Name:",pady=10).grid(row=14,column=0,sticky="E")
            tk.Entry(self.infoFrame, textvariable=self.Fname).grid(row=14,column=1,sticky='w')
            tk.Label(self.infoFrame, text="SD card:",pady=10).grid(row=15,column=0,sticky="E")
            tk.Entry(self.infoFrame, textvariable=self.SD).grid(row=15,column=1,sticky='w')
            tk.Label(self.infoFrame, text="Print Time:",pady=10).grid(row=16,column=0,sticky="E")
            tk.Entry(self.infoFrame, textvariable=self.Ptime).grid(row=16,column=1,sticky='w')

        elif option == 2:
            self.window.geometry("500x275")
            tk.Label(self.infoFrame, text="Enter Reason:",pady=10).grid(row = 50,column = 0,sticky='e')
            tk.Entry(self.infoFrame, textvariable=self.reasonEntry, width=40).grid(row = 50,column = 1,columnspan = 2,sticky='w')
        elif option == 3:
            self.window.geometry("500x400")
            tk.Label(self.infoFrame, text="Date of original message:",pady=10).grid(row = 52,column = 0,sticky='e')
            tk.Entry(self.infoFrame, textvariable=self.dateEntry1).grid(row = 52,column = 1,sticky='w')
            tk.Label(self.infoFrame, text="Last date to pickup print:",pady=10).grid(row = 53,column = 0,sticky='e')
            tk.Entry(self.infoFrame, textvariable=self.dateEntry2).grid(row = 53,column = 1,sticky='w')
        elif option == 4:
            self.window.geometry("500x275")
            tk.Label(self.infoFrame, text="Cancel Request if \nPatron doesn't respond by:",pady=10).grid(row = 54,column = 0,sticky='e')
            tk.Entry(self.infoFrame, textvariable=self.responseDate).grid(row = 54,column = 1,sticky='w')

        tk.Button(self.infoFrame, text=text, command=function,padx=10).grid(row=80,column=1,sticky="e")
        tk.Button(self.infoFrame, text="Back to Menu", command=self.destroyFrame,padx=10).grid(row=80,column=0)

    def destroyFrame(self):
        self.infoFrame.destroy()
        self.backToMenu()

    def defineNewPatronInfo(self):
        filledRows = self.wks.get_all_values()
        end_row = len(filledRows) + 1
        if self.CourseYN == 0:
            CourseYN = "N"
        else:
            CourseYN = "Y"
        if self.research == 0:
            research = "N"
        else:
            research = "Y"
        if self.OwnC == 0:
            OwnC = "N"
        else:
            OwnC = "Y"
        if self.consent == 0:
            consent = "N"
        else:
            consent = "Y"
        self.wks.update_cell(end_row, 1, self.ticketNumEntry.get())
        self.wks.update_cell(end_row, 2, self.nameEntry.get())
        self.wks.update_cell(end_row, 3, self.emailEntry.get())
        self.wks.update_cell(end_row, 4, self.dateToday.get())
        self.wks.update_cell(end_row, 5, self.StaffInitials.get())
        self.wks.update_cell(end_row, 6, CourseYN)
        self.wks.update_cell(end_row, 7, self.CourseCode.get())
        self.wks.update_cell(end_row, 8, self.affiliation.get())
        self.wks.update_cell(end_row, 9, self.department.get())
        self.wks.update_cell(end_row, 10, research)
        self.wks.update_cell(end_row, 11, OwnC)
        self.wks.update_cell(end_row, 12, consent)
        self.wks.update_cell(end_row, 13, self.handle.get())
        self.wks.update_cell(end_row, 14, self.SD.get())
        self.wks.update_cell(end_row, 15, self.Fname.get())
        self.wks.update_cell(end_row, 16, self.Ptime.get())
        infoLab1 = tk.Label(self.infoFrame, text="Submitted!",fg='green')
        infoLab1.grid(row = 100,column = 0,columnspan = 3)
        infoLab1.update()
        self.infoFrame.after(1000, infoLab1.destroy())
        self.infoFrame.destroy()
        self.backToMenu()

    def readyForPickup(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.rowstr = str(self.row_number)
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Ready for Pickup"
        self.msg += "Subject: " + subject + '\n\n'
        body1 = "Hi " + self.name + ",\n\nGood news! The following requested 3D print job has been printed " \
                                    "successfully:\n\n"
        body1 += "Ticket #: " + self.Ticketnum + "\n\nPlease bring this email and your McMaster ID card with you to " \
                 "the Help Desk in Lyons New Media Centre (Mills Library, 4th floor) to retrieve your item.\n\n"
        body1 += "You will be required to sign for it, so a proxy cannot come to pick this up for you.\n\nWe will " \
                 "hold this item for no more than 30 days from today's date before it is reclaimed and/or recycled.  " \
                 "If you cannot make it into the Centre due to work/being home etc., please let us know and we can " \
                 "arrange to hold onto it until you can make it in.\n\nSincerely,\n\nLyons New Media Centre Staff\n\n"
        body1 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body1
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if self.z == "1":
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtreadypickup)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!",fg="green")
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!",fg='green')
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def DelayedPrinting(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Delayed Printing"
        self.msg += "Subject: " + subject + '\n\n'
        body2 = "Hi " + self.name + ",\n\n"
        body2 += "This is in regards to 3D Print Ticket#: " + self.Ticketnum + ".\n\n"
        body2 += "We have had an unusual amount of course-related print requests submitted this term, " \
                 "and are prioritizing those requests before regular requests. Because of this, we may not be able " \
                 "to complete your request by the end of April (last day of exams), so it may be completed as we're " \
                 "going into May (during the summer months).\n\n"
        body2 += "We need to know if you would still like this ticket to be printed, knowing that there is a delay " \
                 "that it may not be printed before the term is over. If you still want it to be printed, but you " \
                 "are not able to pick it up immediately since it may be completed during the summer, you can let us " \
                 "know to hold it for you till you can.\n\n"
        body2 += "Please respond to this email by " + str(self.responseDate.get()) + ". If we do not hear from you " \
                 "by that date, we will assume it is unwanted and will cancel the request.\n\nThank you\n\nLyons " \
                 "New Media Centre\n\n"
        body2 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body2
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if self.z == "1":
            self.wks.update_cell(self.row_number, 17, "Y")
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!",fg='green')
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!",fg='green')
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def Denied(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Denied"
        self.msg += "Subject: " + subject + '\n\n'
        body3 = "Hi " + self.name + ",\n\n"
        body3 += "This is in regards to 3D Print Ticket#: " + self.Ticketnum + ".\n\n"
        body3 += "We are sorry but the following print request has been denied - Ticket#: " + self.Ticketnum + "\n\n"
        body3 += "The reasoning: " + str(self.reasonEntry.get()) + \
                 "\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
        body3 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body3
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if self.z == "1":
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtdenied)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!",fg='green')
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!",fg='green')
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def Clarification_Skewed(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Dimensions Clarifications Needed"
        self.msg += "Subject: " + subject + '\n\n'
        body4 = "Hi " + self.name + ",\n\n"
        body4 += "I'm looking at your print request - Ticket#: " + self.Ticketnum + ".\n\n"
        body4 += "Unfortunately, the dimensions you have submitted appear to skew the 3D model. " \
                 "You can use Cura, a free software to double check your dimensions." \
                 "\n\nOnce you have double checked, feel free to simply reply to this email with the " \
                 "new dimensions in this format:\n\n"
        body4 += "Width (x):\n\nDepth (y):\n\nHeight (z):\n\n"
        body4 += "Once we receive your clarification we will reprocess the print for you!\n\nSincerely,\n\nThe Lyons " \
                 "New Media Centre Staff\n\n "
        body4 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body4
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if self.z == "1":
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtclarification)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!",fg='green')
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!",fg='green')
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def Clarification_Large(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.rowstr = str(self.row_number)

        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Dimensions Clarifications Needed"
        self.msg += "Subject: " + subject + '\n\n'
        body5 = "Hi " + self.name + ",\n\n"
        body5 += "I'm looking at your print request - Ticket#: " + self.Ticketnum + ".\n\n"
        body5 += "Unfortunately, the dimensions you have submitted are too large for our 3D printers to handle. " \
                 "You can use Cura, a free software to resize your model to a size that fits within 5-6 hours.\n\n" \
                 " Once you have double checked the dimensions feel free to simply reply to this email with the new " \
                 "dimensions in this format:\n\n"
        body5 += "Width (x):\n\nDepth (y):\n\nHeight (z):\n\n"
        body5 += "Once we receive your clarification we will reprocess the print for you!\n\nSincerely,\n\nThe" \
                 " Lyons New Media Centre Staff\n\n"
        body5 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body5
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if self.z == "1":
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtclarification)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!",fg='green')
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!",fg='green')
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def Reminder(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.initialDate = str(self.dateEntry1.get())
        self.lastDate = str(self.dateEntry2.get())
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Reminder"
        self.msg += "Subject: " + subject + '\n\n'
        body6 = "Hi " + self.name + ",\n\n"
        body6 += "This is a reminder that your 3D print job - Ticket#:" + self.Ticketnum + " is ready for pickup\n\n"
        body6 += "Please see the original message sent on " + self.initialDate + \
                 " with instructions on picking up the item. If the item is not picked up by " + \
                 self.lastDate + ", we will discard it.\n\n"
        body6 += "If you cannot make it into the Centre due to work/being home etc., please let us know so we can " \
                 "arrange to hold onto it until you can make it in.\n\nSincerely,\n\nThe Lyons New Media Centre " \
                 "Staff\n\n"
        body6 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body6
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if self.z == "1":
            self.wks.update_cell(self.row_number, 17, "Y")
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!",fg="green")
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!",fg='green')
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def Failed(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.rowstr = str(self.row_number)
        else:
            self.name = self.nameEntry.get()
            self.patron_email = self.emailEntry.get()
        self.msg = "Content-Type: text/plain\nMIME-Version: 1.0\n"
        subject = "3D Print Request - Failed"
        self.msg += "Subject: " + subject + '\n\n'
        body7 = "Hi " + self.name + ",\n\n"
        body7 += "We are sorry but the following print request has not printed properly - Ticket#: " + self.Ticketnum +\
                 "\n\n"
        body7 += "What happened / suggestions for printing: " + str(self.reasonEntry.get()) \
                 + "\n\nSincerely,\n\nThe Lyons New Media Centre Staff\n\n"
        body7 += "-- \n\nLyons New Media Centre\n\n4th Floor, Mills Library\n\n"
        self.msg += body7
        LNMC = """library.mcmaster.ca/spaces/lyons"""
        self.msg += LNMC
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(self.User, self.Password)
        server.sendmail(self.User, self.patron_email, self.msg)
        server.quit()
        if self.z == "1":
            self.wks.update_cell(self.row_number, 17, "Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtfailed)
            infoLab2 = tk.Label(self.infoFrame, text="Message Sent & Spreadsheet Updated!",fg='green')
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Message Sent!",fg='green')
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def pickedUp(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.rowstr = str(self.row_number)
            dateToday = date.today().strftime("%m/%d/%Y")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtpickedup)
            self.wks.update_cell(self.row_number, 18, dateToday)
            infoLab2 = tk.Label(self.infoFrame, text="Spreadsheet updated",fg='green')
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Ticket not found, unable to update spreadsheet",fg="red")
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def nevPickedUp(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.rowstr = str(self.row_number)
            self.wks.update_cell(self.row_number, 19, "Reminder email sent but never picked up")
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtneverpickedup)
            infoLab2 = tk.Label(self.infoFrame, text="Spreadsheet updated",fg='green')
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Ticket not found, unable to update spreadsheet",fg='red')
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def cancelled(self):
        if self.z == "1":
            self.Ticketnum = str(self.ticketNumEntry.get())
            self.row_number = self.wks.find(self.Ticketnum).row
            self.name = self.wks.cell(self.row_number, 2).value
            self.patron_email = self.wks.cell(self.row_number, 3).value
            self.rowstr = str(self.row_number)
            self.wks.update_cell(self.row_number, 19, str(self.reasonEntry.get()))
            format_cell_range(self.wks, 'A' + self.rowstr + ':AC' + self.rowstr, self.fmtcancelled)
            print("3D Print has been cancelled")
            print("Spreadsheet Updated")
            infoLab2 = tk.Label(self.infoFrame, text="Spreadsheet updated",fg='green')
            infoLab2.grid(row = 100,column = 0,columnspan = 3)
            infoLab2.update()
            self.infoFrame.after(1000, infoLab2.destroy())
            self.infoFrame.destroy()
            self.backToMenu()
        else:
            infoLab1 = tk.Label(self.infoFrame, text="Ticket not found, unable to update spreadsheet",fg='red')
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            self.infoFrame.after(1000, infoLab1.destroy())
            self.infoFrame.destroy()
            self.backToMenu()

    def findTicket(self, function, text, title, option):
        self.Ticketnum = str(self.ticketNumEntry.get())
        # Exception Handling for when there's no match
        try:
            self.row_number = self.wks.find(self.Ticketnum).row
        except Exception as e:
            self.z = '0'
            infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually",fg = "red")
            infoLab1.grid(row = 100,column = 0,columnspan = 3)
            infoLab1.update()
            infoLab1.after(1000, infoLab1.destroy())
        else:
            if self.Ticketnum == '':
                self.z = '0'
                infoLab1 = tk.Label(self.infoFrame, text="No matching Ticket Number found, Enter Patron info manually",fg='red')
                infoLab1.grid(row = 100,column = 0,columnspan = 3)
                infoLab1.update()
                infoLab1.after(1000, infoLab1.destroy())
            else:
                self.z = '1'
                self.name = self.wks.cell(self.row_number, 2).value
                self.patron_email = self.wks.cell(self.row_number, 3).value
                self.infoFrame.destroy()
                self.getInfo(function, text, title, option)
                infoLab1 = tk.Label(self.infoFrame, text="Found Matching Ticket Number",fg='green')
                infoLab1.grid(row = 100,column = 0,columnspan = 3)
                infoLab1.update()
                infoLab1.after(1000, infoLab1.destroy())


Window()
