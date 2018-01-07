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
from collections import defaultdict

# Email Imports
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# pynYNAB imports
from pynYNAB.Client import nYnabClient
from pynYNAB.connection import nYnabConnection, NYnabConnectionError
#from pynYNAB.schema.budget import Payee, Transaction

import settings

class BudgetLoader:
    """
    Class to simpily loading and organizing our ynab data
    """
    def __init__(self, ynab_user, ynab_password, ynab_budget_name):
        # Set constants
        self.__save_file = 'balances.p'
        
        connection = nYnabConnection(ynab_user, ynab_password)
        connection.init_session()
        self.__client = nYnabClient(nynabconnection=connection, budgetname=ynab_budget_name)

        self.__balances = self.__load_new_balances()
        self.__old_balances = self.__load_old_balances()
        self.__categories, self.__subcategories = self.__get_categories_and_subcategories()

    def __load_old_balances(self):
        if not os.path.isfile(self.__save_file):
            return defaultdict(DefaultBalance)

        old_balances = pickle.load(open(self.__save_file, "rb"))
        if not isinstance(old_balances, defaultdict):
            old_balances = defaultdict(DefaultBalance, old_balances)

        return old_balances

    def __load_new_balances(self):
        """
        Gets current month budget calculations
        """
        balances = defaultdict(DefaultBalance)
        current_year_month = datetime.datetime.now().strftime('%Y-%m')

        for calc in self.__client.budget.be_monthly_subcategory_budget_calculations:
            calc_year_month = calc.entities_monthly_subcategory_budget_id[4:11]
            calc_id = calc.entities_monthly_subcategory_budget_id[12:]

            if calc_year_month == current_year_month:
                balances[calc_id] = calc
                #print(b.entities_monthly_subcategory_budget_id[12:]+': ' + str(b.balance))

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
        message = ''
        
        if diff > 0:
            message += "&nbsp;&nbsp;<span style='color:green'>$" + str(diff) + "&nbsp;&uarr;</span>"
        elif diff < 0:
            message += "&nbsp;&nbsp;<span style='color:red'>$" + str(abs(diff)) + "&nbsp;&darr;</span>"

        return message

    def get_email_message(self):
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
                new_balance = self.__balances[sub_id].balance
                old_balance = self.__old_balances[sub_id].balance
                diff = new_balance - old_balance

                message += '&nbsp;&nbsp;&nbsp;&nbsp;'+ sub_name + ': ' + str(new_balance)
                message += self.__get_styled_diff_string(diff)
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
    server.login(login,password)
    problems = server.sendmail(from_address, to_address_list, msg.as_string())
    server.quit()
    return problems

class DefaultBalance():
    """
    Used as default values for old balances when the value doesn't exist in the dictionary.
    """
    balance = 0


def main():
    print('Getting YNAB info')

    try:
        loader = BudgetLoader(settings.YNAB_USER, settings.YNAB_PASSWORD, settings.YNAB_BUDGET_NAME)
        message = loader.get_email_message()
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
