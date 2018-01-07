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
from pynYNAB.connection import nYnabConnection
#from pynYNAB.schema.budget import Payee, Transaction

import settings

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

class Bal():

    balance = 0


def main():
    print('Getting YNAB info')

    connection = nYnabConnection(settings.YNAB_USER, settings.YNAB_PASSWORD)
    connection.init_session()
    client = nYnabClient(nynabconnection=connection, budgetname=settings.YNAB_BUDGET_NAME)

    cats = {}
    subs = {}
    balances = defaultdict(Bal)

    if os.path.isfile('balances.p'):
        old_balances = pickle.load( open( "balances.p", "rb" ) )
        if type(old_balances) is not defaultdict:
            old_balances = defaultdict(Bal, old_balances)
    else:
        old_balances = defaultdict(Bal)

    #Creates hiarichy structure of category/subcategory and only those that have the keyword in YNAB subcategory notes section
    for cat in client.budget.be_master_categories:
            cats[cat.name]=cat
            subs[cat.name+'_subs'] = {}
            for subcat in client.budget.be_subcategories:
                    if subcat.entities_master_category_id == cat.id:
                            subs[cat.name+'_subs'][subcat.name] = subcat

    #Gets current month budget calculations
    for b in client.budget.be_monthly_subcategory_budget_calculations:
            if b.entities_monthly_subcategory_budget_id[4:11]==(datetime.datetime.now().strftime('%Y-%m')):
                    balances[b.entities_monthly_subcategory_budget_id[12:]]=b
                    #print(b.entities_monthly_subcategory_budget_id[12:]+': ' + str(b.balance))

    #Displays the balance for each subcategory in the subs dict
    bal_str = '<p>'
    for cat in cats:
        if 'Internal' not in cat:
            if len(subs[cat+'_subs'])>0:
                    bal_str += '<b>'+cat+'</b> <br>'
                    for scat in subs[cat+"_subs"]:
                            #print(cat + ' - ' + scat)
                            bal_str += '&nbsp;&nbsp;&nbsp;&nbsp;'+ scat + ': ' + str(balances[subs[cat+"_subs"][scat].id].balance)
                            bal_diff = balances[subs[cat+"_subs"][scat].id].balance - old_balances[subs[cat+"_subs"][scat].id].balance
                            bal_diff = round(bal_diff,2)
                            if bal_diff:
                                if bal_diff > 0:
                                    #Balance goes up
                                    bal_str += "&nbsp;&nbsp;<span style='color:green'>$" + str(bal_diff) + "&nbsp;&uarr;</span>"
                                else:
                                    #Balance went down
                                    bal_str += "&nbsp;&nbsp;<span style='color:red'>$" + str(abs(bal_diff)) + "&nbsp;&darr;</span>"
                            bal_str += '<br>'

    print('Sending Email')

    send_email(
        settings.FROM_ADDRESS,
        settings.TO_LIST,
        'YNAB Balances for ' + datetime.datetime.now().strftime('%x'),
        bal_str,
        settings.GMAIL_USER,
        settings.GMAIL_PASSWORD,
        'smtp.gmail.com:587')

    print('Saving balances')

    pickle.dump( balances, open( "balances.p", "wb" ) )

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        quit()
