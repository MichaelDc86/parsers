import imaplib
import email.message
import email
import email.message
from imapclient import imap_utf7
from bs4 import BeautifulSoup
from pandas import DataFrame, ExcelWriter
import traceback
import lxml


def get_pass():
    server = 'imap.yandex.ru'
    # login = input('Input your email login: example@yandex.ru\n')
    login = "login@yandex.ru"
    # шифрование пароля -------------------------------------------------
    # pwhash = '4a5b9cfbf4b8fefa28d04961405d7c502265e015'
    # password = getpass.getpass("IMAP Password: ")
    #
    # if hashlib.sha1(bytearray(password, 'utf-8')).hexdigest() != pwhash:
    #     print("Invalid password", file=sys.stderr)
    #     sys.exit(1)
    # -------------------------------------------------------------------
    password = ''
    # password = input('Input your password: \n')
    return password, server, login


def get_calls(imap, folder_name_call):
    date_list_call = []
    datas_list_call = []
    unread_messages_call = 0
    message_number_call = []
    status_1, select_data_1 = imap.select(folder_name_call)
    letters_list = imap.search(None, 'ALL')

    for i in letters_list[1][0].split():
        cur_letter = imap.fetch(i, '(RFC822)')
        print('current letter: {}'.format(cur_letter))
        status, data = imap.fetch(i, '(RFC822)')
        message = email.message_from_bytes(data[0][1], _class=email.message.EmailMessage)

        # DATE ---------------------------------------------------------
        date = list(email.utils.parsedate_tz(message['Date']))[:3]
        date_out = ''
        for _ in date:
            date_out += str(_) + '.'
        date_final = date_out[0:-1]
        # --------------------------------------------------------------

        # soup = BeautifulSoup(message.get_payload(decode=True), 'html5lib')
        soup = BeautifulSoup(message.get_payload(decode=True), 'lxml')

        try:

            # DATAS --------------------------
            datas = soup.findAll('p')

            print(datas)
            datas_out = list(map(lambda x: x.text, datas))
            print(datas_out)
            datas_final = ''
            for _ in datas_out:
                if _:
                    datas_final += _ + '  '  # '\n' + '  '
            datas_final = datas_final[0:]
            print(datas_final)
            # --------------------------------------------------------------

            # DATAFRAME ---------------------------------------------------------

            date_list_call.append(date_final)
            datas_list_call.append(datas_final)

            # -------------------------------------------------------------------

            print(date_list_call, datas_list_call)

        except IndexError as err:
            print('Ошибка:\n', traceback.format_exc())
            print('IndexError')
            unread_messages_call += 1
            message_number_call.append(i)

    return date_list_call, datas_list_call, unread_messages_call, message_number_call


def main():
    global folder_name
    global folder_name_call
    unread_messages = 0
    unread_messages_call = 0
    message_number = []
    message_number_call = []
    date_list = []
    shipment_list = []
    address_list = []
    datas_list = []
    price_list = []
    date_list_call = []
    datas_list_call = []

    # CONNECT -----------------------------------------------
    password, server, login = get_pass()
    print("Connecting to {}...".format(server))
    imap = imaplib.IMAP4_SSL(server)
    print("Connected! Logging in as {}...".format(login))
    imap.login(login, password)
    print("Logged in! Listing folders...")
    # --------------------------------------------------------
    folders = []
    for item in imap.list()[1]:
        folder_raw = imap_utf7.decode(item)
        try:
            folder = folder_raw.split('"')[3]
        except IndexError:
            folder = folder_raw.split('"')[2]

        if folder == 'INBOX|юмакс заказы':
            folder_number = imap.list()[1].index(item)
            folder_name = imap.list()[1][folder_number].decode('utf-8').split('"')[-2]

        if folder == 'юммаксзвонки':
            print('Haaaaalllooooo!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
            # print(item, type(item))
            # print(str(item).split('"'))
            folder_number_call = imap.list()[1].index(item)
            folder_name_call = imap.list()[1][folder_number_call].decode('utf-8').split('"')[-2]

        folders.append(folder)
        print(folder)
    # ENTER THE FOLDER ------------------------------------------------

    date_list_call, datas_list_call, unread_messages_call, message_number_call = get_calls(imap, folder_name_call)

    status, select_data = imap.select(folder_name)
    letters_list = imap.search(None, 'ALL')

    # ----------------------------------------------------------------
    # CYCLE FOR LETTERS ----------------------------------------------
    # for i in ['26']:
    # for i in [str(x) for x in range(1, 5)]:

    for i in letters_list[1][0].split():
        cur_letter = imap.fetch(i, '(RFC822)')
        print('current letter: {}'.format(cur_letter))
        status, data = imap.fetch(i, '(RFC822)')  # .encoding('utf-8')   '(RFC822)'
        message = email.message_from_bytes(data[0][1], _class=email.message.EmailMessage)

        # DATE ---------------------------------------------------------
        date = list(email.utils.parsedate_tz(message['Date']))[:3]
        date_out = ''
        for _ in date:
            date_out += str(_) + '.'
        date_final = date_out[0:-1]
        # --------------------------------------------------------------

        # soup = BeautifulSoup(message.get_payload(decode=True), 'html5lib')
        soup = BeautifulSoup(message.get_payload(decode=True), 'lxml')

        try:
            # PRICE -----------------------------------------------------
            try:
                price = 0
                if soup.strong:
                    price = soup.strong.string.split('=')[0]
                    price = [x for x in price if x.isdigit() or x == '.']
                    price_str = ''.join(price)
                    price = float(price_str)
            # ------------------------------------------------------------

                # DATAS --------------------------
                datas = soup.findAll('p')[0].findAll('b')
                print(datas)
                datas_out = list(map(lambda x: x.string, datas))
                print(datas_out)
                datas_final = ''
                for _ in datas_out:
                    if _:
                        datas_final += _ + '  '  # '\n' + '  '
                datas_final = datas_final[0:]

                # --------------------------------------------------------------

                # ADDRESS ------------------------------------------------------
                address = str(soup.findAll('p')[2].contents[0])
                print(address)
                if ',' in list(address):
                    address = str(address).replace(',', '')

                # --------------------------------------------------------------
                shipment = ''
                if len(soup.findAll('p')) > 4:
                    shipment = str(soup.findAll('p')[4].contents[0])
                    shipment = shipment.replace('\\', '')
                    print(shipment)

                # DATAFRAME ---------------------------------------------------------

                date_list.append(date_final)
                shipment_list.append(shipment)
                price_list.append(price)
                address_list.append(address)
                datas_list.append(datas_final)

                # -----------------------------------------------------------------------

                print(date_final, shipment, price, address, datas_final)
            except AttributeError as e:
                print('Ошибка:\n', traceback.format_exc())
                print('AttributeError')
                unread_messages += 1
                message_number.append(i)
        except IndexError as err:
            print('Ошибка:\n', traceback.format_exc())
            print('IndexError')
            unread_messages += 1
            message_number.append(i)

    print('unread messages : {}'.format(unread_messages))
    print('unread messages call: {}'.format(unread_messages_call))
    print('unread messages numbers: {}'.format(message_number))
    print('unread messages numbers call: {}'.format(message_number_call))
    print(datas_list)
    print(datas_list_call)
    writer = ExcelWriter('Заказы и звонки.xlsx')

    df = DataFrame({'Date': date_list, 'Shipment': shipment_list, 'Price': price_list, 'Address': address_list, 'Personal data': datas_list})
    df_call = DataFrame({'Date': date_list_call, 'Personal data': datas_list_call})

    df.to_excel(writer, 'Заказы', index=False)
    df_call.to_excel(writer, 'Звонки', index=False)

    writer.save()
    imap.logout()


if __name__ == '__main__':
    main()
