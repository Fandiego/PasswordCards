import os
import re
import secrets
import openpyxl
import xlsxwriter
import glob

# Diego Fantino
# since 31. October 2021

alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789{%}?!._;-\'[]#'
head = ['AB', 'CD', 'EF', 'GH', 'IJ', 'KL', 'MN', 'OP', 'QR', 'ST', 'UV', 'WX', 'YZ', '.,', '_;', '-:', '#%']
col = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']
row = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20']
underscore = "-------------------------------------------------------" \
             "-------------------------------------------------------"


def userdialog():
    print("Wilkommen")
    user_input = ""

    # While Loop for Dialog
    while user_input.lower() != 'x':
        print("Drücken sie c um eine Passwortkarte zu erstellen")
        print("Drücken sie o um eine Passwortkarte zu öffnen")
        print("Drücken sie d um eine Passwortkarte zu löschen")
        print("Drücken sie x um das Programm zu schliessen")
        user_input = input("Eingabe >")

        # while input is wrong
        while user_input.lower() != 'c' and user_input.lower() != 'o' and \
                user_input.lower() != 'd' and user_input.lower() != 'x':
            print(underscore)
            print("Falsche Eingabe")
            user_input = input("Eingabe >")

        # generate password card
        if user_input.lower() == 'c':
            print(underscore)
            print("Geben sie den Namen der Passwortkarte die sie generieren wollen ein")
            name = input("Eingabe >")

            while os.path.isfile("PasswordCards/" + name + ".xlsx"):
                print(underscore)
                print("Diese Datei existert schon, geben sie bitte einen anderen Namen ein!")
                name = input("Eingabe >")

            generate_pw_card(name)
            load_pw_card(name)

        # open password card
        elif user_input.lower() == 'o':
            print(underscore)
            print_all_pw_cards()
            print(underscore)
            print("Geben sie den Namen der Passwortkarte die sie öffnen wollen ein")
            name = input("Eingabe >")

            while os.path.isfile("PasswordCards/" + name + ".xlsx") is False:
                print(underscore)
                print("Diese Datei existert noch nicht, geben sie bitte einen anderen Namen ein!")
                name = input("Eingabe >")

            load_pw_card(name)

        # delete password card
        elif user_input.lower() == 'd':
            print("In Progress")

        else:
            pass


# generates password card
def generate_pw_card(file):
    # get the Path of the File
    path = "PasswordCards/" + file + ".xlsx"

    # open the File from Path
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    # generate password card
    for i in range(17):
        for k in range(19):
            worksheet.write(col[i] + row[k + 1], random_string())
        worksheet.write(col[i] + row[0], head[i])

    # close File
    workbook.close()


# loads password card
def load_pw_card(file):
    # Get Path of File
    path = "PasswordCards/" + file + ".xlsx"

    # Open File from Path
    wb_object = openpyxl.load_workbook(path)
    sheet = wb_object.active

    # get Keyword
    keyword = get_keyword()

    # get Keyword length
    keyword_length = len(keyword)

    colnumb = 1
    password = ''

    # get Password from Keyword
    for i in range(keyword_length):
        for j in range(17):
            ch = head[j]
            if ch[0].lower() == keyword[i].lower() or ch[1].lower() == keyword[i].lower():
                password = password + sheet[col[j] + row[colnumb]].value
                if colnumb != 19:
                    colnumb = colnumb + 1
                else:
                    colnumb = 1

    # Print Password
    print("Ihr Passwort ist: " + password)

    # Close File
    wb_object.close()
    print(underscore)


# deletes password card
def delete_pw_card(file):
    print("in progress " + file)
    print(underscore)


# prints all password cards
def print_all_pw_cards():
    print("PasswordCards: ")
    for file in glob.glob("PasswordCards/" + "*.xlsx"):
        file = file.split("\\")
        file = file[1].split(".")
        print(file[0])


# returns a random String
def random_string():
    return secrets.choice(alphabet) + secrets.choice(alphabet) + secrets.choice(alphabet) + secrets.choice(
        alphabet)


# gets the keyword from user input
def get_keyword():
    regmatch = False
    keyword = ""
    print(underscore)
    print("Geben sie hier Ihr Schlüsselwort ein")
    while regmatch is False:
        keyword = input("Schlüsselwort >")
        if re.match(r'[a-zA-Z9.,_;\-:#%]', keyword):
            regmatch = True
        else:
            print(underscore)
            print("Schlüsselwort enthält nicht erlaubte Zeichen")

    return keyword


userdialog()
