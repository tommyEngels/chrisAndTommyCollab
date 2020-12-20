
import smtplib
import pandas as pd

# EMAIL_ADDRESS =  os.environ.get('EMAIL_USER')
# EMAIL_ADDRESS =  os.environ.get('EMAIL_PASS')

'''
TODO: Import excel 
TODO: read sheet and desirable columns
TODO: send 1 email per order number
    for loop for each order number
        for loop for respective line items for said order
            store relevant data into format
        send 1 email per order including all line items to broker
    

'''
# Read Excel workbook and relevant page
# order_report = pd.ExcelFile("/Users/sideHustle/Desktop/Freelance/sampleSavor/python_scripts/Order Fulfillment Workbook.xlsm")
# df1 = pd.read_excel(xls, '8-27-2020')
#

order_report = pd.read_excel("orderReport.xlsx")

all_order_numbers = order_report['Name']
all_broker_emails = order_report['Email']
all_broker_names = order_report['Shipping Name']
all_line_items = order_report['Lineitem name']
all_shipping_names = order_report['Shipping Company']

for i in range(len(order_report)):

    order_number = str(all_order_numbers[i])
    broker_email = str(all_broker_emails[i])
    shipping_name = str(all_shipping_names[i])
    line_items = ''
    while all_order_numbers[i] == all_order_numbers[i+1]:
        line_items = all_line_items[i] + '. \n'
        i = i + 1   
    # TODO: send 1 email per order number
    print("email sent to "+broker_email)
    print("deliver to "+shipping_name)
    print("order includes: "+line_items)
    

        

# with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
#     smtp.ehlo() #identifies ourselves with mailserver
#     smtp.starttls() #encrypts traffic
#     smtp.ehlo()#identified as encrypted traffic

#     smtp.login("person1coder@gmail.com", "Testuser1!")

#     subject = 'Dinner tonight?'
#     body = '6pm at Mcdonalds?'

#     msg = f'Subject: {subject}\n\n{body}'

#     smtp.sendmail("person1coder@gmail.com", "tbengels@gmail.com", msg)





