import win32com.client
from datetime import datetime
from pywintypes import com_error
from db import Database
from openpyxl import Workbook

database = Database()

class App:

    def __init__(self):
        self.olapp = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
        self.index_list = []
        self.chosen_mailbox = None
        self.box = self.olapp.Folders
        self.fldrs = None
        self.chosen_choice = None
        self.current_mailbox = None
        self.export_list = []
        self.resolved_name_ind = ["solv", "SOLV", "ompleted", "OMPLETED", "RAITEE", "raitee", "ompletado", "OMPLETADO"]
        self.folder_string = "        // Total:[{}], Unread:[{}], Read:[{}], Modified Today: [{}]"
        self.received_today = 0
        self.resolved_today = 0
        self.start_menu()

    def start_menu(self):
        self.index_list = []
        choices = ["List mailboxes", "Enter mailbox name", "Exit"]
        print("***********")
        print("*MAIN MENU*")
        print("***********")
        for ind, itm in enumerate(choices):
            print("[", ind, "]", itm)
            self.index_list.append(ind)
        self.start_menu_choice()

    def start_menu_choice(self):
        chosen_choice = None
        try:
            chosen_choice = int(input("Your choice: "))
        except ValueError:
            self.start_menu()
        if chosen_choice not in self.index_list:
            print("wrong choice:")
            self.start_menu()
        elif chosen_choice == 0:
            self.list_mailboxes()
        elif chosen_choice == 1:
            self.enter_mailbox()
        elif chosen_choice == 2:
            exit()

    def enter_mailbox(self):
        chosen_mailbox = input("Enter mailbox name: ")
        self.export_list.clear()
        headings_list = ["No.", "Folder Name", "Total Items", "Read", "Modified Today"]
        self.export_list.append(headings_list)
        try:
            self.fldrs = self.box(chosen_mailbox).Folders
            snt = self.fldrs("Sent Items")
            sent_today = self.folder_stats(snt)

            self.received_today = 0  # resetting received, and resolved
            self.resolved_today = 0

            inb = self.fldrs("Inbox")
            for ind in self.resolved_name_ind:
                if ind in inb.Name:
                    self.resolved_today += self.count_resolved(inb)
                    break

            exp_slist = ["", inb.Name, self.count_items(inb)[0], self.count_items(inb)[1], self.count_items(inb)[2],
                         self.count_items(inb)[3]]

            print(exp_slist[1], self.folder_string.format(exp_slist[2], exp_slist[3], exp_slist[4], exp_slist[5]))
            self.export_list.append(exp_slist)
            self.received_today += self.folder_stats(inb)

            self.inspect_folder(inb, 1)

            print("")
            print("")
            print("Today's stats:")
            print("Emails received today: ", self.received_today)
            print("Emails sent today: ", sent_today)
            print("Emails resolved today: ", self.resolved_today)
            n = datetime.now()
            database.insert(self.box(chosen_mailbox).Name, self.received_today, sent_today, self.resolved_today,
                            n.year, n.month, n.day, n.isocalendar()[1])
            self.enter_mailbox_menu()
        except com_error:
            print("")
            print("")
            print("ERROR! Wrong/unavailable mailbox name! If the mailbox name is correct, please restart outlook!")
            print("")
            print("")
            self.enter_mailbox_menu()

    def inspect_folder(self, folder, level):
        folder_string = "        // Total:[{}], Unread:[{}], Read:[{}], Modified Today: [{}]"
        es = "    "
        def_str = "{} [{}] {} {}"
        for index, item in enumerate(folder.Folders):
            for ind in self.resolved_name_ind:
                if ind in item.Name:
                    self.resolved_today += self.count_resolved(item)
                    break
            exp_slist = [index, item.Name, self.count_items(item)[0], self.count_items(item)[1],
                         self.count_items(item)[2], self.count_items(item)[3]]

            print(def_str.format(level * es, index, item.Name, folder_string.format(exp_slist[2],
                                                                                    exp_slist[3],
                                                                                    exp_slist[4],
                                                                                    exp_slist[5])))
            self.export_list.append(exp_slist)
            self.received_today += self.folder_stats(item)

            self.inspect_folder(item, level+1)  # circular reference to create infinite loop



    def list_mailboxes(self):
        for ind, folder in enumerate(self.box):
            if ("RS" in folder.Name or "@" in folder.Name) and "Archive" not in folder.Name:
                print("[", ind, "]", folder.Name)
        self.start_menu()

    def enter_mailbox_menu(self):
        submenu_choices = ['Reports', 'To Excel', 'Main Menu']
        print("***********")
        print("*MAILBOX MENU*")
        print("***********")
        index_list = []
        for ind, itm in enumerate(submenu_choices):
            print("[", ind, "]", itm)
            index_list.append(ind)
        choice = None
        choice = int(input("Your choice: "))

        if choice not in index_list:
            self.enter_mailbox_menu()
        elif choice == 0:
            self.view_all()
        elif choice == 1:
            self.to_excel()
        elif choice == 2:
            self.start_menu()

    @staticmethod
    def count_items(ol_item):
        today = datetime.now().strftime("%Y-%m-%d") + " 00:00"
        ol_item_items = ol_item.Items
        ol_item_items_rst = ol_item_items.Restrict("[UnRead] = True")
        ol_item_items_reso1 = ol_item_items.Restrict("[LastModificationTime] > '" + today + "'")
        return_list = [ol_item_items.Count, ol_item_items_rst.Count, ol_item_items.Count - ol_item_items_rst.Count,
                       ol_item_items_reso1.Count]

        return return_list

    @staticmethod
    def count_resolved(ol_item):
        today = datetime.now().strftime("%Y-%m-%d") + " 00:00"
        ol_item_items = ol_item.Items
        ol_item_items_reso1 = ol_item_items.Restrict("[LastModificationTime] > '" + today + "'")

        return ol_item_items_reso1.Count

    @staticmethod
    def view_all():
        headers = ("id", "mailbox", "received", "sent", "resolved", "year", "month", "day", "week")
        print(headers)
        for row in database.view_all():
            print(row)

    @staticmethod
    def folder_stats(folder):
        today = datetime.now().strftime("%Y-%m-%d") + " 00:00"
        if folder.Name == "Sent Items":
            sent_items = folder.Items
            sent_today = sent_items.Restrict("[SentOn] > '" + today + "'")
            return sent_today.Count
        else:
            all_items = folder.Items
            all_items_r = all_items.Restrict("[ReceivedTime] >= '" + today + "'")
            return all_items_r.Count

    def to_excel(self):
        wb = Workbook()
        ws = wb.active
        for index, item in enumerate(self.export_list):
            for subindex, subitem in enumerate(item):
                ws.cell(row=index+1, column=subindex+1).value = subitem
                wb.save("export.xlsx")
        self.enter_mailbox_menu()


if __name__ == '__main__':
    app = App()





