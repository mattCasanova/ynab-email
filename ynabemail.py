#!/usr/bin/python3.5

"""
Grabs budget information (names and balances) and then sends an email to a list of email addresses.
Useful if you want to keep someone updated on category balances but they won't regularly
check the application.
"""
# Standard imports
import datetime
import pickle
import os.path

# Email Imports
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# pynYNAB imports
from pynYNAB.Client import nYnabClient
from pynYNAB.connection import nYnabConnection, NYnabConnectionError

import settings

class BudgetLoader:
    """
    Class to simpily loading and organizing our ynab data
    """
    def __init__(self, ynab_user, ynab_password, ynab_budget_name):
        # Set constants
        self.__save_file = 'balances.p'
        self.__money_format = '${:,.2f}'

        connection = nYnabConnection(ynab_user, ynab_password)
        connection.init_session()
        self.__client = nYnabClient(nynabconnection=connection, budgetname=ynab_budget_name)
        self.__client.sync()

        self.__balances = self.__load_new_balances()
        self.__old_balances = self.__load_old_balances()
        self.__categories, self.__subcategories = self.__get_categories_and_subcategories()

    def __load_old_balances(self):
        if not os.path.isfile(self.__save_file):
            return dict()

        old_balances = pickle.load(open(self.__save_file, "rb"))
        return old_balances

    def __load_new_balances(self):
        """
        Gets current month budget calculations
        """
        balances = dict()
        current_year_month = datetime.datetime.now().strftime('%Y-%m')

        for calc in self.__client.budget.be_monthly_subcategory_budget_calculations:
            calc_year_month = calc.entities_monthly_subcategory_budget_id[4:11]
            calc_id = calc.entities_monthly_subcategory_budget_id[12:]

            if calc_year_month == current_year_month:
                balances[calc_id] = calc

        return balances

    def __get_categories_and_subcategories(self):
        """
        Creates hiarichy structure of category/subcategory and only those that
        have the keyword in YNAB subcategory notes section
        """

        categories = {}
        subcategories = {}

        for category in self.__client.budget.be_master_categories:
            categories[category.name] = category
            subcategories[category.name] = {}

            for subcategory in self.__client.budget.be_subcategories:
                if subcategory.entities_master_category_id == category.id:
                    subcategories[category.name][subcategory.name] = subcategory

        return categories, subcategories

    def __get_styled_diff_string(self, diff):
        if diff == 0:
            return ''

        color = 'green'
        direction = '&uarr;'
        message = "&nbsp;&nbsp;<span style='color:{}'>{}&nbsp;{}</span>"

        if diff < 0:
            color = 'red'
            direction = '&darr;'

        return message.format(color, self.__money_format.format(diff), direction)

    def create_email_body(self):
        """
        Displays the balance for each subcategory in an html message
        """

        message = '<p>'
        for cat_name in self.__categories:
            if 'Internal' in cat_name or not self.__subcategories[cat_name]:
                continue

            message += '<b>' + cat_name + '</b> <br>'
            for sub_name in self.__subcategories[cat_name]:

                sub_id = self.__subcategories[cat_name][sub_name].id

                obj = self.__old_balances.get(sub_id)
                old_balance = obj.balance if obj else 0
                new_balance = self.__balances.get(sub_id).balance

                message += '&nbsp;&nbsp;&nbsp;&nbsp;'
                message += sub_name + ': '
                message += self.__money_format.format(new_balance)
                message += self.__get_styled_diff_string(new_balance - old_balance)
                message += '<br>'

        return message

    def save_balances(self):
        """
        Saves current month balances to a file
        """
        pickle.dump(self.__balances, open(self.__save_file, "wb"))

def send_email(from_address, to_address_list, subject, message, login, password, smtpserver):
    """
    Sends the given email message to the list of addresses via an smpt server
    """

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = from_address
    msg['To'] = ','.join(to_address_list)

    htmlmessage = '<html><head></head><body>'+message+'</body></html>'

    part1 = MIMEText(message, 'plain')
    part2 = MIMEText(htmlmessage, 'html')

    msg.attach(part1)
    msg.attach(part2)

    server = smtplib.SMTP(smtpserver)
    server.starttls()
    server.login(login, password)
    problems = server.sendmail(from_address, to_address_list, msg.as_string())
    server.quit()
    return problems

def main():
    print('Getting YNAB info')

    try:
        loader = BudgetLoader(settings.YNAB_USER, settings.YNAB_PASSWORD, settings.YNAB_BUDGET_NAME)
        message = loader.create_email_body()
    except NYnabConnectionError:
        return

    print('Sending Email')

    send_email(
        settings.FROM_ADDRESS,
        settings.TO_LIST,
        'YNAB Balances for ' + datetime.datetime.now().strftime('%x'),
        message,
        settings.GMAIL_USER,
        settings.GMAIL_PASSWORD,
        'smtp.gmail.com:587')

    print('Saving balances')
    loader.save_balances()

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        quit()
