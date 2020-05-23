import datetime
import operator
import smtplib
import ssl
import xlrd
import os
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup
from selenium import webdriver
from time import sleep
from xlsxwriter import Workbook


class App:
    def __init__(self, username='username', password='password', target_city='city', stay=7,
                 path='//home/user/Booking'):
        self.username = username
        self.password = password
        self.target_city = target_city
        self.stay = stay
        self.path = path
        self.driver = webdriver.Firefox(
            executable_path='/usr/local/bin/geckodriver')  # Change this to your FirefoxDriver path.
        self.error = False
        self.main_url = 'http://www.restel.es'
        self.all_hotels = []
        self.all_prices = []
        self.all_addresses = []
        self.display = []
        self.cheap = []
        self.index = ""
        self.options = {}
        self.driver.get(self.main_url)
        sleep(1)
        self.log_in()
        if self.error is False:
            # todo: REF:
            #  https://es.stackoverflow.com/questions/109086/esperar-respuestas-para-continuar-selenium-python
            # The explicit wait, unlike an implicit one or what time.sleep does (although this is blocking)
            # however it was the solution for this search engine (compare with solole project where at this
            # point sleep was useless)
            sleep(1)  #### fixme: explicit wait
            self.search_target_profile()
        if self.error is False:
            self.scroll_down()
        if self.error is False:
            if not os.path.exists(path):
                os.mkdir(path)
            self.file_manager()
        print('\nsending email...\n')
        sleep(10)
        self.driver.close()

    def log_in(self, ):
        try:
            # fixme: the xpath did not work selected with the right mouse button (neither with the css selector). The
            #  solution has been to use the variable "placeholder" as xpath, taking it from the same script used in
            #  solole.
            user_name_input = self.driver.find_element_by_xpath('//form/div[1]/div/input')
            user_name_input.send_keys(self.username)
            sleep(1)

            password_input = self.driver.find_element_by_xpath('//form/div[2]/div/input')
            password_input.send_keys(self.password)
            sleep(1)

            user_name_input.submit()
            sleep(1)

        except Exception:
            print('Some exception occurred while trying to find username or password field')
            self.error = True

    def flip_calendar(self, days):
        today = datetime.datetime.utcnow()
        print("check in:  ", today)
        check_out = today + datetime.timedelta(days - 1)
        print("check out: ", check_out)

        flip = check_out.month - today.month
        return flip

    def search_target_profile(self):
        try:
            search_bar = self.driver.find_element_by_css_selector('#filterHotels')
            search_bar.send_keys(self.target_city)
            # fixme: WARNING: immediately after entering the city in the field, in this case, we need an
            #  explicit wait of at least one second before clicking to display correctly the drop-down menu:
            sleep(2)
            search_bar.click()
            # enter destination city
            target_city = self.driver.find_element_by_css_selector(
                "li.item:nth-child(1) > div:nth-child(2) > span:nth-child(1)")
            target_city.click()
            sleep(1)

            # calendar picker
            self.driver.find_element_by_css_selector('#calendarHotels').click()
            sleep(1)
            if self.flip_calendar(self.stay) == 0:
                # todo: accessing a drop-down calendar item by position within the list
                #  https://selenium-python.readthedocs.io/navigating.html#interacting-with-the-page
                all_options = self.driver.find_elements_by_class_name('available')
                all_options[0].click()
                all_options = self.driver.find_elements_by_class_name('available')
                all_options[self.stay - 1].click()
                sleep(2)
            else:
                all_options = self.driver.find_elements_by_class_name('available')
                all_options[0].click()
                self.driver.find_element_by_css_selector('div.drp-calendar:nth-child(3)').click()

            # search button
            login_button = self.driver.find_element_by_xpath('//*[@id="search-hotels"]')
            # instead of submit it works with click
            login_button.click()
            sleep(3)
        except Exception:
            self.error = True
            print('Could not find search bar')

    def scroll_down(self):
        global position
        self.driver.implicitly_wait(20)

        soup = BeautifulSoup(self.driver.page_source, 'lxml')
        hotel_list = soup.find_all('div', {'class': 'element'})
        euro_symbol = '€'

        # todo REF: https://stackoverflow.com/questions/48006078/how-to-scroll-down-in-python-selenium-step-by-step
        # FIXME 1: two ways to scroll down,
        #  1) go down to the bottom of the page at once.
        # self.driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        # FIXME 2:
        #  2) Descend from item to item to the bottom of the page.
        # in this example and item is the text of the button "See options":
        read_mores = self.driver.find_elements_by_xpath('//a[text()="Ver opciones"]')
        for read_more in read_mores:
            self.driver.execute_script("arguments[0].scrollIntoView();", read_more)
            # read_more.click()

        print("\n\tdisplay:\n")
        try:
            self.driver.implicitly_wait(20)
            for i, hotel in enumerate(hotel_list):

                hotel_price = hotel.find('span', {'class': 'final-price'}).getText().strip('€')
                hotel_price = hotel_price.replace('.', '')
                hotel_price = hotel_price.replace(',', '.')
                hotel_price = float(hotel_price)
                hotel_price = "{0:.2f}".format(hotel_price)
                self.all_prices.append(hotel_price)
                if len(hotel_price) == 5:
                    hotel_price = "   " + hotel_price
                if len(hotel_price) == 6:
                    hotel_price = "  " + hotel_price
                if len(hotel_price) == 7:
                    hotel_price = " " + hotel_price

                hotel_name = hotel.find('a', {'class': 'hotel-name'}).getText()
                hotel_address = hotel.find('span', {'class': 'address-content'}).getText()
                self.all_hotels.append(hotel_name)
                self.all_addresses.append(hotel_address)
                if i < 9:
                    print(" %d - %s %s %s - %s" % (i + 1, hotel_price, euro_symbol, hotel_name, hotel_address))
                else:
                    print("%d - %s %s %s - %s" % (i + 1, hotel_price, euro_symbol, hotel_name, hotel_address))

            print("\n\tranking:\n")
            # float cast
            new_prices_2 = []
            for element in self.all_prices:
                rank = float(element)
                new_prices_2.append(rank)

            # final list
            # dict version allow just 2 values 'k' and 'v':
            # list = dict(zip(self.all_hotels, new_prices_2))
            # ranking_2 = sorted(list.items(), key=operator.itemgetter(1))
            # for k, v in ranking_2:...etc'''
            # with common zip package we easily pack the addresses too:
            display_list = list(zip(self.all_hotels, new_prices_2, self.all_addresses))
            ranking_2 = sorted(display_list, key=operator.itemgetter(1))
            # todo REF: https://discuss.codecademy.com/t/how-can-i-sort-a-zipped-object/454412/6
            for k, v, w in ranking_2:
                if v < 100.00:
                    print("   ", "{0:.2f}".format(v), k)
                if 99.00 < v < 1000.00:
                    print("  ", "{0:.2f}".format(v), k)
                if 999.00 < v < 10000.00:
                    print(" ", "{0:.2f}".format(v), k)
                if v > 9999.00:
                    print("", "{0:.2f}".format(v), k)

            self.cheap = ranking_2[0]
            self.options = ranking_2
            print('\ncheapest reservations: ', self.cheap[0], self.cheap[1], euro_symbol)
            # self.display = display_list[7]
            # print('Target button number: ', self.display.index(self.cheap[0]))
            self.display = display_list
            for i, collation in enumerate(display_list):
                if collation[0] == self.cheap[0]:
                    position = i
            print('position of the target button: ', position + 1)
            self.index = str(position - 1)
            if self.error is False:
                self.target_button(self.index)

            sleep(2)
        except Exception as e:
            self.error = True
            print(e)
            print('Some error occurred while trying to scroll down')

    def target_button(self, index):
        target_button = self.driver.find_element_by_xpath(
            '//app-search-results-list/div/div[1]/div/div[1]/div[' + index + ']/div/div[3]/div/div[3]/a')
        self.driver.execute_script("arguments[0].scrollIntoView();", target_button)
        # target_button.click()

    def file_manager(self, ):
        bookings_folder_path = os.path.join(self.path, 'bookings')
        if not os.path.exists(bookings_folder_path):
            os.mkdir(bookings_folder_path)
        if self.error is False:
            self.write_bookings_to_excel_file(bookings_folder_path)
        if self.error is False:
            self.read_bookings_from_excel_file(self.path + '/bookings/bookings.xlsx')

    def read_bookings_from_excel_file(self, excel_path):
        workbook = xlrd.open_workbook(excel_path)
        worksheet = workbook.sheet_by_index(0)
        for row in range(2):
            col_1, col_2, col_3, col_4, col_5, col_6 = worksheet.row_values(row)
            print(col_1, '    ', col_2, '    ', )

    def write_bookings_to_excel_file(self, booking_path):
        print('\nwriting to excel...')
        workbook = Workbook(os.path.join(booking_path, 'bookings.xlsx'))
        worksheet = workbook.add_worksheet()
        worksheet.set_column(2, 3, 50)
        worksheet.set_column(1, 1, 9)
        worksheet.set_column(4, 4, 9)
        bold = workbook.add_format({'bold': True})
        cell_format = workbook.add_format({'bold': True, 'italic': True, 'font_color': 'blue'})
        cell_money = workbook.add_format({'bold': True, 'italic': True, 'font_color': 'blue', 'num_format': '#,##0.00'})
        money = workbook.add_format({'num_format': '#,##0.00'})
        row = 0
        worksheet.write(row, 0, 'Code', bold)  # 3 --> row number, column number, value
        worksheet.write(row, 1, 'Price', bold)
        worksheet.write(row, 2, 'Hotel', bold)
        worksheet.write(row, 3, 'Address', bold)
        worksheet.write(row, 4, 'Retail', bold)
        worksheet.write(row, 5, 'Profit', bold)

        row += 1

        worksheet.write(row, 0, 'BEST', cell_money)
        worksheet.write(row, 1, self.cheap[1], cell_format)
        worksheet.write(row, 2, self.cheap[0], cell_format)
        worksheet.write(row, 3, self.cheap[2], cell_format)
        worksheet.write_formula(1, 4, '=1.374*B2', cell_money)
        worksheet.write_formula(1, 5, '=E2-B2', cell_money)
        row += 1

        for i, option in enumerate(self.options):
            if i < 9:
                worksheet.write(row, 0, 'AA0' + str(i + 1))
            else:
                worksheet.write(row, 0, 'AA' + str(i + 1))
            worksheet.write(row, 1, option[1], money)
            worksheet.write(row, 2, option[0])
            worksheet.write(row, 3, option[2])
            worksheet.write_array_formula('E3:E31', '{=1.374*B3:B31}', money)
            worksheet.write_array_formula('F3:F31', '{=E3:E31-B3:B31}', money)
            row += 1
        workbook.close()
        # fixme WARNING:
        # in order for xlsxwriter to create the spreadsheet, the workbook must be closed
        # right at the end. If it closes after it won't create it, check it by closing it after the line:
        # self.send_attachment (spreadsheet)
        spreadsheet = '//home/jmartorell/Booking/bookings/bookings.xlsx'
        self.send_attachment(spreadsheet)

    def send_attachment(self, file):
        subject = "An email with attachment from Python"
        body = "This is an email with attachment sent from Python"
        sender_email = "SENDER'S EMAIL ACCOUNT"
        receiver_email = "RECEIVER'S EMAIL ACCOUNT"
        # password = input("Type your password and press enter:")
        password = 'ZXspectrum5128$}_'

        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        message["Bcc"] = receiver_email  # Recommended for mass emails

        # Add body to email
        message.attach(MIMEText(body, "plain"))

        filename = file  # In same directory as script

        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()

        # Log in to server using secure context and send email
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, text)


if __name__ == '__main__':
    app = App()