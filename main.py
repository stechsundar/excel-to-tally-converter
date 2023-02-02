import dataclasses
import streamlit as st
import pandas as pd
import requests
import openpyxl
import xml.etree.ElementTree as ET
from time import sleep

st.set_page_config(page_title="Excel to Tally", page_icon=":guardsman:", layout="wide",
                   initial_sidebar_state="collapsed")
newcmplist = []

newurl = "http://localhost:9000"
global svcurrentcompany
xmldata = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST><TYPE>COLLECTION</TYPE><ID>ListOfCompanies</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION Name='ListOfCompanies'><TYPE>Company</TYPE><FETCH>Name,CompanyNumber</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"


def get_company_names(xmldata):
    i = 0
    page = ''
    while page == '':
        try:
            page = requests.get(newurl, data=xmldata)
            break
        except ConnectionError:
            st.write("Connection refused by the server..")
            st.write("Let me sleep for 5 seconds")
            st.write("ZZzzzz...")
            sleep(5)
            st.write("Was a nice sleep, now let me continue...")
            continue
    root = ET.fromstring(page.text.strip())
    for cmp in root.findall('./BODY/DATA/COLLECTION/COMPANY'):
        cmp_name = cmp.get('NAME')
        newcmplist.append(cmp_name)
        i += 1

    return newcmplist


newcmplist = get_company_names(xmldata)
svcurrentcompany = st.selectbox("Select the Company", newcmplist)


@st.cache
def load_data(file_path):
    data = pd.read_excel(file_path).fillna("")
    data['DATE'] = pd.to_datetime(data['DATE'])
    data['DATE'] = data['DATE'].dt.strftime('%d-%m-%Y')

    number_columns = ['AMT', 'QTY']

    # format the number columns with 2 decimal places

    data[number_columns] = data[number_columns].round(2)
    data.index += 1
    return data


def payentry(vrdt, area, ledname, amt, narr):
    newurl = "http://localhost:9000"

    try:
        if amt == 0:
            return

        cramt = amt
        dritemvaluestr = "-" + str(amt)
        critemvaluestr = str(cramt)
        new_data = '<ENVELOPE><HEADER><VERSION> 1 </VERSION><TALLYREQUEST>Import</TALLYREQUEST><TYPE>Data</TYPE>'
        new_data += '<ID>Vouchers</ID></HEADER><BODY><DESC><STATICVARIABLES>'
        new_data += '<SVCURRENTCOMPANY>' + str(svcurrentcompany).rstrip() + '</SVCURRENTCOMPANY>'
        new_data += '</STATICVARIABLES></DESC><DATA>'
        new_data += '<TALLYMESSAGE xmlns:UDF="TallyUDF">'
        new_data += '<VOUCHER VCHTYPE="SALES" ACTION="Create" OBJVIEW = "InvoiceVoucherView">'
        new_data += '<DATE>' + vrdt + '</DATE>'
        new_data += '<NARRATION>' + narr + "(" + area + ")" + '</NARRATION>'
        new_data += '<VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>'
        new_data += '<PERSISTEDVIEW>Account Voucher View</PERSISTEDVIEW>'
        new_data += '<PARTYNAME>' + ledname + '</PARTYNAME>'
        new_data += '<PARTYLEDGERNAME>' + ledname + '</PARTYLEDGERNAME>'
        new_data += '<LEDGERENTRIES.LIST>'
        new_data += '<LEDGERNAME>' + ledname + '</LEDGERNAME>'
        new_data += '<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'
        new_data += '<AMOUNT>' + dritemvaluestr + '</AMOUNT>'
        new_data += '</LEDGERENTRIES.LIST>'
        new_data += '<LEDGERENTRIES.LIST>'
        new_data += '<LEDGERNAME>Cash</LEDGERNAME>'
        new_data += '<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>'
        new_data += '<AMOUNT>' + critemvaluestr + '</AMOUNT>'
        new_data += '</LEDGERENTRIES.LIST>'
        new_data += '</VOUCHER>'
        new_data += '</TALLYMESSAGE>'
        new_data += '</DATA></BODY></ENVELOPE>'

        page = ''
        while page == '':
            try:
                page = requests.post(url=newurl, data=new_data)
                break
            except ConnectionError:
                print("Connection refused by the server..")
                print("Let me sleep for 5 seconds")
                print("ZZzzzz...")
                sleep(5)
                print("Was a nice sleep, now let me continue...")
                continue

    except Exception as e:
        st.error(f"An error occured: {e}")


def recentry(vrdt, area, ledname, amt, narr):
    newurl = "http://localhost:9000"

    try:
        if amt == 0:
            return

        cramt = amt
        dritemvaluestr = "-" + str(amt)
        critemvaluestr = str(cramt)
        new_data = '<ENVELOPE><HEADER><VERSION> 1 </VERSION><TALLYREQUEST>Import</TALLYREQUEST><TYPE>Data</TYPE>'
        new_data += '<ID>Vouchers</ID></HEADER><BODY><DESC><STATICVARIABLES>'
        new_data += '<SVCURRENTCOMPANY>' + str(svcurrentcompany).rstrip() + '</SVCURRENTCOMPANY>'
        new_data += '</STATICVARIABLES></DESC><DATA>'
        new_data += '<TALLYMESSAGE xmlns:UDF="TallyUDF">'
        new_data += '<VOUCHER VCHTYPE="SALES" ACTION="Create" OBJVIEW = "InvoiceVoucherView">'
        new_data += '<DATE>' + vrdt + '</DATE>'
        new_data += '<NARRATION>' + narr + "(" + area + ")" + '</NARRATION>'
        new_data += '<VOUCHERTYPENAME>Receipt</VOUCHERTYPENAME>'
        new_data += '<PERSISTEDVIEW>Account Voucher View</PERSISTEDVIEW>'
        new_data += '<PARTYNAME>' + ledname + '</PARTYNAME>'
        new_data += '<PARTYLEDGERNAME>' + ledname + '</PARTYLEDGERNAME>'
        new_data += '<LEDGERENTRIES.LIST>'
        new_data += '<LEDGERNAME>Cash</LEDGERNAME>'
        new_data += '<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'
        new_data += '<AMOUNT>' + dritemvaluestr + '</AMOUNT>'
        new_data += '</LEDGERENTRIES.LIST>'
        new_data += '<LEDGERENTRIES.LIST>'
        new_data += '<LEDGERNAME>' + ledname + '</LEDGERNAME>'
        new_data += '<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>'
        new_data += '<AMOUNT>' + critemvaluestr + '</AMOUNT>'
        new_data += '</LEDGERENTRIES.LIST>'
        new_data += '</VOUCHER>'
        new_data += '</TALLYMESSAGE>'
        new_data += '</DATA></BODY></ENVELOPE>'

        page = ''
        while page == '':
            try:
                page = requests.post(url=newurl, data=new_data)
                break
            except ConnectionError:
                print("Connection refused by the server..")
                print("Let me sleep for 5 seconds")
                print("ZZzzzz...")
                sleep(5)
                print("Was a nice sleep, now let me continue...")
                continue
    except Exception as e:
        st.error(f"An error occured: {e}")


def pur_entry(vrdt, area, itemname, itemunit, qty, ratevar, ledname, narr, amt):
    newurl = "http://localhost:9000"

    try:
        if qty == 0:
            return

        cramt = amt
        dritemvaluestr = "-" + str(amt)
        critemvaluestr = str(cramt)
        itemqtystr = str(qty) + " " + itemunit
        itemratestr = str(ratevar)
        new_data = '<ENVELOPE><HEADER><VERSION> 1 </VERSION><TALLYREQUEST>Import</TALLYREQUEST><TYPE>Data</TYPE>'
        new_data += '<ID>Vouchers</ID></HEADER><BODY><DESC><STATICVARIABLES>'
        new_data += '<SVCURRENTCOMPANY>' + str(svcurrentcompany).rstrip() + '</SVCURRENTCOMPANY>'
        new_data += '</STATICVARIABLES></DESC><DATA>'
        new_data += '<TALLYMESSAGE xmlns:UDF="TallyUDF">'
        new_data += '<VOUCHER VCHTYPE="PURCHASE" ACTION="Create" OBJVIEW = "InvoiceVoucherView">'
        new_data += '<DATE>' + vrdt + '</DATE>'
        new_data += '<NARRATION>' + narr + "(" + area + ")" + '</NARRATION>'
        new_data += '<VOUCHERTYPENAME>PURCHASE</VOUCHERTYPENAME>'
        new_data += '<ISINVOICE>Yes</ISINVOICE>'
        new_data += '<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>'
        new_data += '<PARTYNAME>' + ledname + '</PARTYNAME>'
        new_data += '<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>'
        new_data += '<PARTYLEDGERNAME>' + ledname + '</PARTYLEDGERNAME>'
        new_data += '<LEDGERENTRIES.LIST>'
        new_data += '<LEDGERNAME>' + ledname + '</LEDGERNAME>'
        new_data += '<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>'
        new_data += '<AMOUNT>' + critemvaluestr + '</AMOUNT>'
        new_data += '</LEDGERENTRIES.LIST>'
        new_data += '<ALLINVENTORYENTRIES.LIST>'
        new_data += '<STOCKITEMNAME>' + str(itemname).rstrip() + '</STOCKITEMNAME>'
        new_data += '<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'
        new_data += '<ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>'
        new_data += '<RATE>' + itemratestr + '</RATE>'
        new_data += '<AMOUNT>' + dritemvaluestr + '</AMOUNT>'
        new_data += '<ACTUALQTY>' + itemqtystr + '</ACTUALQTY>'
        new_data += '<BILLEDQTY>' + itemqtystr + '</BILLEDQTY>'
        new_data += '<ACCOUNTINGALLOCATIONS.LIST>'
        new_data += '<LEDGERNAME>PURCHASE</LEDGERNAME>'
        new_data += '<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'
        new_data += '<AMOUNT>' + dritemvaluestr + '</AMOUNT>'
        new_data += '</ACCOUNTINGALLOCATIONS.LIST>'
        new_data += '</ALLINVENTORYENTRIES.LIST>'
        new_data += '</VOUCHER>'
        new_data += '</TALLYMESSAGE>'
        new_data += '</DATA></BODY></ENVELOPE>'

        page = ''
        while page == '':
            try:
                page = requests.post(url=newurl, data=new_data)
                break
            except ConnectionError:
                print("Connection refused by the server..")
                print("Let me sleep for 5 seconds")
                print("ZZzzzz...")
                sleep(5)
                print("Was a nice sleep, now let me continue...")
                continue
    except Exception as e:
        st.error(f"An error occured: {e}")


def sales_entry(vrdt, area, itemname, itemunit, qty, ratevar, ledname, narr, amt):
    newurl = "http://localhost:9000"

    try:
        if qty == 0:
            return

        cramt = amt
        dritemvaluestr = "-" + str(amt)
        critemvaluestr = str(cramt)
        itemqtystr = str(qty) + " " + itemunit
        itemratestr = str(ratevar)
        new_data = '<ENVELOPE><HEADER><VERSION> 1 </VERSION><TALLYREQUEST>Import</TALLYREQUEST><TYPE>Data</TYPE>'
        new_data += '<ID>Vouchers</ID></HEADER><BODY><DESC><STATICVARIABLES>'
        new_data += '<SVCURRENTCOMPANY>' + str(svcurrentcompany).rstrip() + '</SVCURRENTCOMPANY>'
        new_data += '</STATICVARIABLES></DESC><DATA>'
        new_data += '<TALLYMESSAGE xmlns:UDF="TallyUDF">'
        new_data += '<VOUCHER VCHTYPE="SALES" ACTION="Create" OBJVIEW = "InvoiceVoucherView">'
        new_data += '<DATE>' + vrdt + '</DATE>'
        new_data += '<NARRATION>' + narr + "(" + area + ")" + '</NARRATION>'
        new_data += '<VOUCHERTYPENAME>SALES</VOUCHERTYPENAME>'
        new_data += '<ISINVOICE>Yes</ISINVOICE>'
        new_data += '<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>'
        new_data += '<PARTYNAME>' + ledname + '</PARTYNAME>'
        new_data += '<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>'
        new_data += '<PARTYLEDGERNAME>' + ledname + '</PARTYLEDGERNAME>'
        new_data += '<LEDGERENTRIES.LIST>'
        new_data += '<LEDGERNAME>' + ledname + '</LEDGERNAME>'
        new_data += '<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'
        new_data += '<AMOUNT>' + dritemvaluestr + '</AMOUNT>'
        new_data += '</LEDGERENTRIES.LIST>'
        new_data += '<ALLINVENTORYENTRIES.LIST>'
        new_data += '<STOCKITEMNAME>' + str(itemname).rstrip() + '</STOCKITEMNAME>'
        new_data += '<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>'
        new_data += '<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>'
        new_data += '<RATE>' + itemratestr + '</RATE>'
        new_data += '<AMOUNT>' + critemvaluestr + '</AMOUNT>'
        new_data += '<ACTUALQTY>' + itemqtystr + '</ACTUALQTY>'
        new_data += '<BILLEDQTY>' + itemqtystr + '</BILLEDQTY>'
        new_data += '<ACCOUNTINGALLOCATIONS.LIST>'
        new_data += '<LEDGERNAME>SALES</LEDGERNAME>'
        new_data += '<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>'
        new_data += '<AMOUNT>' + critemvaluestr + '</AMOUNT>'
        new_data += '</ACCOUNTINGALLOCATIONS.LIST>'
        new_data += '</ALLINVENTORYENTRIES.LIST>'
        new_data += '</VOUCHER>'
        new_data += '</TALLYMESSAGE>'
        new_data += '</DATA></BODY></ENVELOPE>'

        page = ''
        while page == '':
            try:
                page = requests.post(url=newurl, data=new_data)
                break
            except ConnectionError:
                print("Connection refused by the server..")
                print("Let me sleep for 5 seconds")
                print("ZZzzzz...")
                sleep(5)
                print("Was a nice sleep, now let me continue...")
                continue
    except Exception as e:
        st.error(f"An error occured: {e}")


def color_negative_red(val):
    if isinstance(val, str):
        return ''
    color = 'red' if val < 0 else 'black'
    return 'color: %s' % color


def main():
    file_path = st.file_uploader("Upload an Excel file", type="xlsx")
    if file_path is not None:
        data = load_data(file_path)
        st.dataframe(data)

        if st.button("Pass Data to Tally"):

            with st.spinner("Passing vouchers to Tally..."):
                sales = 0
                purc = 0
                count = 0
                pymnt = 0
                recvou = 0
                for index, row in data.iterrows():
                    vrdt = row[0]
                    vrtype = row[1].rstrip()
                    rec_pay = row[2].rstrip()
                    area = row[3].rstrip()
                    ratevar = row[4]
                    itemunit = row[5].rstrip()
                    itemname = row[6].rstrip()
                    ledname = row[7].rstrip()
                    narr = row[8]
                    amt = row[9]
                    qty = row[10]
                    if amt == '' and qty == '':
                        continue
                    if rec_pay == "RECEIPTS" and vrtype == "SALES" and qty != "":
                        sales += 1
                        count += 1
                        sales_entry(vrdt, area, itemname, itemunit, qty, ratevar, ledname, narr, amt)
                    if rec_pay == "RECEIPTS" and vrtype == "SALES" and amt != "":
                        recvou += 1
                        count += 1
                        recentry(vrdt, area, ledname, amt, narr)
                    if rec_pay == "PAYMENT" and vrtype == "PURCHASE" and qty != "":
                        purc += 1
                        count += 1
                        pur_entry(vrdt, area, itemname, itemunit, qty, ratevar, ledname, narr, amt)
                    if rec_pay == "PAYMENT" and vrtype != "PURCHASE" and amt != "":
                        pymnt += 1
                        count += 1
                        payentry(vrdt, area, ledname, amt, narr)

                st.write("Total Vouchers : " + str(count))
                st.write("Sales Vouchers : " + str(sales))
                st.write("Purchase Vouchers: " + str(purc))
                st.write("Payment Vouchers: " + str(pymnt))
                st.write("Receipt Vouchers: " + str(recvou))

                st.success("All the vouchers passed to TallyPrime Successfully")


if __name__ == "__main__":
    main()
