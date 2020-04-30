import csv
from woocommerce import API
from datetime import datetime
from tkinter import *


def website_switch(website, websiteid, woo_key, woo_secret, auth_allow):
    wcapi = API(
        url=website,
        consumer_key=woo_key,
        consumer_secret=woo_secret,
        wp_api=True,
        version="wc/v3",
        query_string_auth=auth_allow,
        timeout=60
    )

    website_id = websiteid

    return wcapi, website_id

websites = [
    {Enter information based on wcapi data}
]

def printtext():
    global e
    global string
    string = e.get()
    root.destroy()


root = Tk()
root.title('Change Export File Name')
root.geometry("300x200")
e = Entry(root)
e.focus_set()
var = StringVar()
label = Label(root, textvariable=var)
b = Button(root, width=10, height=2, text="Submit", command=printtext)
var.set("What is the file name extension? For example '-1'")
label.pack(side=TOP, padx=5, pady=10)
e.pack(side=TOP, ipadx=5, ipady=5, padx=5, pady=5)
b.pack(side=TOP, padx=5, pady=5)
root.mainloop()

fileName = 'enter filepath that you want to save your information'.format(string)

with open(fileName, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(["Date", "Time", "Time Zone", "Name", "Type", "Status", "Gross", "Fee", "Net", "From Email Address", "To Email Address", "Transaction ID", "Counterparty Status", "Address Status", "Item Title", "Item ID", "Shipping and Handling Amount", "Insurance Amount", "Sales Tax", "Option 1 Name", "Option 1 Value", "Option 2 Name", "Option 2 Value", "Auction Site", "Buyer ID", "Item URL", "Closing Date", "Escrow Id", "Invoice Id", "Reference Txn ID", "Invoice Number", "Ship Via", "Quantity", "Receipt ID", "Balance", "Address Line 1", "Address Line 2/District/Neighborhood", "Town/City", "State/Province/Region/County/Territory/Prefecture/Republic", "Zip/Postal Code", "Country", "Contact Phone Number", "Different Address?"])
    
    for website in range(len(websites)):
        web_name = websites[website]['website']
        web_id = websites[website]['websiteid']
        web_key = websites[website]['woo_key']
        web_secret = websites[website]['woo_secret']
        offset = websites[website]['offset']
        auth_allow = websites[website]['auth_allow']
        
        wcapi, website_id = website_switch(web_name, web_id, web_key, web_secret, auth_allow)

        orders = 100

        date = datetime.strptime("Time that you want to use")
        iso_date = date.isoformat()

        orders = wcapi.get("orders", params={"per_page": orders, "status": "processing", "after": iso_date, "offset": int(offset)}).json()

        for i in range(jnorders):  # First loop for orders
            
            try:
                order = orders[i]
            except IndexError:
                break

            fees = order['fee_lines']

            fee_price = 0
            for fee in range(len(fees)):
                fee_price += float(fees[fee]['total'])

            ship_via = order['shipping_lines'][0]['method_title']

            def shipping_classes(method):
                switcher = {
                    "Enter a shipping method"
                }
                return switcher.get(method, "FE")

            date_created = str(order['date_created'])[:10]
            first_name = str(order['shipping']['first_name']).replace("’", "").replace("'", "") # order first name
            last_name = str(order['shipping']['last_name']).replace("’", "").replace("'", "") # order last name
            email = str(order['billing']['email']) # order last name
            total = str(order['total']) # order total
            company = str(order['shipping']['company']).replace("’", "").replace("'", "") # order company
            shipping_total = str(order['shipping_total'])  # order shipping_tax
            total_tax = str(order['total_tax']) # order total_tax
            number = str(order['number']) # order number
            address_1 = str(order['shipping']['address_1']).replace("’", "").replace("'", "") # order address_1
            address_2 = str(order['shipping']['address_2']).replace("’", "").replace("'", "") # order address_2
            city = str(order['shipping']['city']) # order city
            state = str(order['shipping']['state']) # order state
            postcode = str(order['shipping']['postcode']) # order postcode
            country = str(order['shipping']['country']) # order country

            substitutions = [
                ("-", ""), 
                ("?", ""),
                ("+1", ""),
                (".", ""),
                ("(", ""),
                (")", ""),
                (" ", "")
            ]

            phone = str(order['billing']['phone']) # order phone number

            for search, replacment in substitutions:
                phone = phone.replace(search, replacment)

            payment_method_title = "TEL: {} | {}".format(phone, str(order['payment_method_title'])[:6]) # order payment_method_title

            writer.writerow([
                date_created,  # Column A (Date)
                '', '', # Empty Column B & C (Time & Time Zone)
                '{} {}'.format(first_name, last_name), # Column D (Name)
                'Authorization', # Column E (Type)
                '', # Empty Column F (Status)
                total,  # Column G (Gross)
                '', # Empty Column H (Fee)
                total, # Column I (Net)
                "Enter email address that you want to have",  # Column J (From Email Address)
                "Enter email address that you want to have",  # Column K (To Email Address)
                company, # Column L (Transaction ID)
                '', # Empty Column M (Counterparty Status)
                '', # Empty Column N (Address Status)
                '', # Empty Column O (Item Title)
                '',  # Empty Column P (Item ID)
                shipping_total, # Empty Column Q (Shipping and Handling Amount)
                '',  # Empty Column R (Insurance Amount)
                total_tax, # Column S (Sales Tax)
                '',  # Empty Column T (Option 1 Name)
                '',  # Empty Column U (Option 1 Value)
                '',  # Empty Column V (Option 2 Name)
                '',  # Empty Column W (Option 2 Value)
                '',  # Empty Column X (Auction Site)
                '',  # Empty Column Y (Buyer ID)
                '',  # Empty Column Z (Item URL)
                '',  # Empty Column AA (Closing Date)
                '',  # Empty Column AB (Escrow Id)
                '',  # Empty Column AC (Invoice Id)
                '',  # Empty Column AD (Reference Txn ID)
                '{}-{}'.format(website_id, number), # Column AE (Invoice Number)
                shipping_classes(str(ship_via)),  # Column AF (Ship Via)
                '',  # Empty Column AG (Quantity)
                '',  # Empty Column AH (Receipt ID)
                '',  # Empty Column AI (Balance)
                address_1, # Column AJ (Address Line 1)            
                address_2, # Column AK (Address Line 2/District/Neighborhood)
                city, # Column AL (Town/City)
                state, # Column AM (State/Province/Region/County/Territory/Prefecture/Republic)
                postcode, # Column AN (Zip/Postal Code)
                country, # Column AO (Country)
                phone, # Column AP (Contact Phone Number)
                payment_method_title, # Column AQ (Different Address?)
            ])

            items = order['line_items'] # line items array 

            for li in range(len(items)): # Nested loop for line items

                lineItem = order['line_items'][li] # line item index
                
                line_item_total = lineItem['total']
                name = lineItem['name'].replace("’", "").replace("'", "")  # line item title
                sku = lineItem['sku'] # line item sku
                quantity = lineItem['quantity'] # line item quantity

                writer.writerow([
                    date_created,  # Column A (Date)
                    '', '',  # Empty Column B & C (Time & Time Zone)
                    '{} {}'.format(first_name, last_name),  # Column D (Name)
                    'Shopping Cart Item',  # Column E (Type)
                    '',  # Empty Column F (Status)
                    line_item_total,  # Column G (Gross)
                    '',  # Empty Column H (Fee)
                    '',  # Empty Column I (Net)
                    email, # Column J (From Email Address)
                    "Enter email address that you want to have", # Column K (To Email Address)
                    company, # Column L (Transaction ID)
                    '',  # Empty Column M (Counterparty Status)
                    '',  # Empty Column N (Address Status)
                    name,  # Column O (Item Title)
                    sku,  # Column P (Item ID)
                    '', # Empty Column Q (Shipping and Handling Amount)
                    '',  # Empty Column R (Insurance Amount)
                    '',  # Column S (Sales Tax)
                    '',  # Empty Column T (Option 1 Name)
                    '',  # Empty Column U (Option 1 Value)
                    '',  # Empty Column V (Option 2 Name)
                    '',  # Empty Column W (Option 2 Value)
                    '',  # Empty Column X (Auction Site)
                    '',  # Empty Column Y (Buyer ID)
                    '',  # Empty Column Z (Item URL)
                    '',  # Empty Column AA (Closing Date)
                    '',  # Empty Column AB (Escrow Id)
                    '', # Empty Column AC (Invoice Id)
                    '',  # Empty Column AD (Reference Txn ID)
                    '{}-{}'.format(website_id, number), # Column AE (Invoice Number)
                    '',  # Empty Column AF (Ship Via)
                    quantity,  # Column AG (Quantity)
                    '',  # Empty Column AH (Receipt ID)
                    '',  # Empty Column AI (Balance)
                    address_1, # Column AJ (Address Line 1)
                    address_2, # Column AK (Address Line 2/District/Neighborhood)
                    city, # Column AL (Town/City)
                    state, # Column AM (State/Province/Region/County/Territory/Prefecture/Republic)
                    postcode, # Column AN (Zip/Postal Code)
                    country, # Column AO (Country)
                    phone, # Column AP (Contact Phone Number)
                    payment_method_title, # Column AQ (Different Address?)
                ])
    
print("complete")    

def close_window_box():
    root_fin.destroy()

root_fin = Tk()
root_fin.title('Finished')
root_fin.geometry("300x200")
var_fin = StringVar()
label_fin = Label(root_fin, textvariable=var_fin)
b_fin = Button(root_fin, width=10, height=2, text="OK", command=close_window_box)
var_fin.set("Complete!")
label_fin.pack(side=TOP, padx=5, pady=10)
b_fin.pack(side=TOP, padx=5, pady=5)
root_fin.mainloop()
