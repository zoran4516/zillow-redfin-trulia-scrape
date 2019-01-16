# -*- coding: utf-8 -*-

__author__ = 'ceho'

import os
import sys
import traceback
import xlsxwriter
import datetime
import time
import lxml
import cssselect

userAgent = 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36'

from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import selenium.webdriver.support.ui as ui

shortIDs = False

def findElement(node, css):
    try:
        if shortIDs:
            parts = css.split('#')
            if len(parts) == 2:
                if parts[1].startswith('ctl00_'):
                    css = parts[0] + '#' + parts[1][6:]
        res = node.find_element_by_css_selector(css)
    except NoSuchElementException:
        res = None
    except:
        res = None
    return res

def getText(node, css, attr):
    res = ''
    elem = findElement(node, css)
    if elem:
        if attr == '':
            res = elem.text
        else:
            res = elem.get_attribute(attr)
        if not res:
            res = ''
    return res

def goToURL(linkName, css):
    try:
        links = driver.find_elements_by_css_selector('table[id*="ContentPlaceHolder1_wizDetails_SideBarContainer_SideBarList"] a')
        for link in links:
            if link.text == linkName:
                link.click()
                wait = ui.WebDriverWait(driver, 30)
                wait.until(lambda driver: findElement(driver, css))
                return True
    except:
        pass

def getURL(url):
    print url
    driver.get(url)
    while True:
        elem = findElement(driver, 'div[id="ctl00_ContentPlaceHolder1_pnlCaptcha"]')
        if elem is None:
            break
        time.sleep(1)

def parseItem(url, writer):
    getURL(url)

    data = []
    data.append(getText(driver, 'span#ctl00_ContentPlaceHolder1_lblPTID', ''))
    data.append(getText(driver, 'textarea#ctl00_ContentPlaceHolder1_wizDetails_txtAddressAddress', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtAddressSubdivision', 'value'))
    data.append(getText(driver, 'textarea#ctl00_ContentPlaceHolder1_wizDetails_txtAddressLegalDescription', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_chkAgriculture', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCurrentOwnerName', 'value'))
    data.append(getText(driver, 'textarea#ctl00_ContentPlaceHolder1_wizDetails_txtCurrentOwnerAddress', 'value'))

    goToURL('Basics', "input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsEADDate")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsEADDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsEADReceptionNumber', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsOriginalSaleDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsActualSaleDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtDateFileReceived', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtDateFileCreated', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsEADRecordingDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsEADRecordingReceptionNumber', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsDOTDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsDOTRecorded', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsDOTReceptionNumber', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsLoanInfoLoanType', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsLoanInfoOriginalAmount', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsLoanInfoCurrentDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsLoanInfoCurrentAmount', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsLoanInfoInterestRate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsLoanInfoInterestType', 'value'))
    data.append(getText(driver, 'textarea#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsLoanInfoCurrentLender', 'value'))
    data.append(getText(driver, 'textarea#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsLoanInfoGrantee', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtBasicsLoanInfoGrantor', 'value'))

    goToURL('Cure', "input#ctl00_ContentPlaceHolder1_wizDetails_txtInentDeadline")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtInentDeadline', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCureDeadline', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCuredAmountReceived', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCuredBy', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCuredDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCureFiguresRequested', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCureFiguresReceived', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCureFiguresExpire', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCureFiguresTotalToCure', 'value'))

    goToURL('Deed', "input#ctl00_ContentPlaceHolder1_wizDetails_txtDeedDeedToDate")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtDeedDeedToDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtDeedDeedTo', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtDeedRecipentAddress', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtDeededDeedRecptionNumber', 'value'))

    goToURL('Law Firm', "input#ctl00_ContentPlaceHolder1_wizDetails_txtLawFirmName")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLawFirmName', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLawFirmFileNumber', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLawFirmAddress', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLawFirmTelephone', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLawFirmFax', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLawFirmEmail', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLawFirmContact', 'value'))

    goToURL('Mailings', "input#ctl00_ContentPlaceHolder1_wizDetails_txtMailingsInitialCRNoticeMailed")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtMailingsInitialCRNoticeMailed', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtMailingsInitialSaleNoticeMailed', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtMailingsMailed', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtDeferredCombinedNoticeMailed', 'value'))

    if goToURL('Owner Redemption', "input#ctl00_ContentPlaceHolder1_wizDetails_txtOwnerRedemptionRedemptionTime"):
        data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtOwnerRedemptionRedemptionTime', 'value'))
        data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtOwnerRedemptionDateIntentToRedeemTime', 'value'))
        data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtOwnerRedemptionAfterSaleLastDateToRedeem', 'value'))
        data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtOwnerRedemptionExtendedLastDateToRedeem', 'value'))
        data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtOwnerRedemptionRedeemedDate', 'value'))
        data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtOwnerRedemptionRedeemedAmount', 'value'))
        data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtRedemptionPaid', 'value'))
        data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtRedemptionRefund', 'value'))
        data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtOwnerRedemptionCertificateOdRedemptionReceptionNumber', 'value'))
    else:
        data.extend(['', '', '', '', '', '', '', '', ''])

    goToURL('Publication', "input#ctl00_ContentPlaceHolder1_wizDetails_txtPublicationPublication")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtPublicationPublication', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtPublicationFirstPublicationDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtPublicationLastPublicationDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtPublicationFirstRePubDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtPublicationLastRePubDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtFirstDefermentPublicationDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLastDefermentPublicationDate', 'value'))

    goToURL('Redemption', "input#ctl00_ContentPlaceHolder1_wizDetails_txtLienorsNoticeOfIntentToRedeemFilingDeadline")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLienorsNoticeOfIntentToRedeemFilingDeadline', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLienorsRedPerEnds', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtLienorsExtendedLastDateForOwnerToRedeem', 'value'))

    goToURL('Sale Information', "input#ctl00_ContentPlaceHolder1_wizDetails_txtCopPendingBid")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCopPendingBid', 'value'))
    data.append(getText(driver, 'textarea#ctl00_ContentPlaceHolder1_wizDetails_txtCopBidderInformation', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCopDeficiencyAmount', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCopTotalIndebtedness', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCopActualSoldDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCopBidAmount', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtRevisedDeficiencyAmountDuetoOutsideBid', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCopOverbidAmount', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCOPIssuedTo', 'value'))
    data.append(getText(driver, 'textarea#ctl00_ContentPlaceHolder1_wizDetails_txtCOPIssuedToAddress', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCOPAssignedTo', 'value'))
    data.append(getText(driver, 'textarea#ctl00_ContentPlaceHolder1_wizDetails_txtCOPAssignedToAddress', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCopRecordedUnderReceptionNumber', 'value'))

    goToURL('Withdrawal', "input#ctl00_ContentPlaceHolder1_wizDetails_txtWithdrawalToBeWithdrawnDate")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtWithdrawalToBeWithdrawnDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtWithdrawalWithdrawnDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtWithdrawalWithdrawnReceptionNumber', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtWithdrawalVoidRecorded', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtWithdrawalVoidReceptionNumber', 'value'))

    goToURL('Deferments', "input#ctl00_ContentPlaceHolder1_wizDetails_txtAffidavitOfPostingReceivedDate")

    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtAffidavitOfPostingReceivedDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCertificateOfQualificationReceivedDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtApprovedQualificationForDefermentDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtCounselorContactedDate', 'value'))
    data.append(getText(driver, 'input#ctl00_ContentPlaceHolder1_wizDetails_txtDefermentTerminatedDate', 'value'))

    return data

def waitForLoading(driver):
    while True:
        try:
            time.sleep(2)
            elem = findElement(driver, "div#ctl00_ContentPlaceHolder1_up1progress")
            if elem is None:
                break
            style = elem.get_attribute('style')
            if style == 'display: none;':
                break
        except:
            pass

def parsePage(url, fname, date1, date2):
    try:
        _parsePage(url, fname, date1, date2)
    except:
        print "Error: %s" % traceback.format_exc()

def _parsePage(url, fname, date1, date2):
    f = xlsxwriter.Workbook(fname + '.xlsx')
    writer = f.add_worksheet('Sheet1')
    writer.freeze_panes(1, 0)
    writer.autofilter('A1:CP1')
    writer.c = -1
    writerow(writer, ['County', 'ID', 'STATUS', 'PROPERTY ADDRESS', 'Subdivision', 'Legal Description', 'Agricultural', 'CURRENT OWNER', 'Address',
          'NED Date', 'NED Reception', 'Originally Scheduled Sale Date', 'Currently Scheduled Sale Date', 'Date File Received',
          'Date File Created', 'NED RERECORDING', 'Reception #', 'DEED OF TRUST', 'Recorded', 'Reception #', 'Loan Type',
          'Original Principal Balance', 'Principal Balance As Of Date', 'Outstanding Principal Balance', 'Interest Rate',
          'Interest Type', 'Current Holder', 'Grantee (Original Beneficiary)', 'Original Grantor (Borrower)',
          'Intent Deadline', 'Cure Deadline', 'Cured Amount Received', 'Cured By', 'Date Cured', 'Cure Figures Requested',
          'Cure Figures Received', 'Cure Figures Expire', 'Cure Figures Total To Cure', 'Date Deed Recorded', 'Deeded To',
          'Recipient Address', 'Deed Reception #', 'LAW FIRM', 'File Number', 'Address', 'Telephone', 'Fax', 'Email', 'Contact',
          'Initial Rights Mailed', 'Notice of Sale Mailing Date', 'Amended Sale and Rights Notice Mailed',
          'Deferred Combined Notice Mailed', 'Redemption Time', 'Date Intent To Redeem Filed', 'After Sale Last Date to Redeem',
          'Extended Last Date to Redeem', 'Redeemed Date', 'Redemption Amount Due', 'Redemption Amount Paid',
          'Redemption Refund Amount', 'Reception #', 'Published In', 'First Publication Date', 'Last Publication Date',
          'First Re-Pub Date', 'Last Re-Pub Date', 'Deferred Publication Date', 'Last Deferred Publication Date',
          'REDEMPTION', 'All Redemption Periods Expire', 'Extended Last Date to Redeem', 'Holders Initial Bid',
          'Holder', 'Deficiency Amount', 'Total Indebtedness', 'Date Sold', 'Successful Bid at Sale', 'Deficiency Amount Post Sale',
          'Overbid Amount', 'COP Issued To', 'COP Issued To Address', 'COP Assigned To', 'COP Assigned To Address',
          'Reception #', 'To Be Withdrawn Date', 'Withdrawn Date', 'Reception #', 'Withdrawal Recorded Date',
          'Withdrawal Reception number', 'Affidavit of Posting Received Date', 'Certificate of Qualification Received Date',
          'Counselor Approved Homeowner Qualification for Deferment Date', 'Homeowner Contacted Counselor Date',
          'Deferment Terminated Date'])

    driver.get(url)
    count = 0
    while count < 30:
        count += 1
        elem = findElement(driver, 'input[name="ctl00$ContentPlaceHolder1$btnAcceptTerms"]')
        if elem:
            elem.send_keys(Keys.RETURN)
            time.sleep(5)

        elem = findElement(driver, 'input[name="ctl00$ContentPlaceHolder1$txtSoldDate1"]')
        if elem:
            elem.send_keys(date1)

        elem = findElement(driver, 'input[name="ctl00$ContentPlaceHolder1$txtSoldDate2"]')
        if elem:
            elem.send_keys(date2)
            elem.send_keys(Keys.RETURN)
            break

    waitForLoading(driver)

    pag = []
    elem = findElement(driver, 'table.SearchResultsGrid tr')
    if elem:
        pages = elem.find_elements_by_css_selector('table tr td a')
        for page in pages:
            if not page.text.isdigit():
                continue
            j = page.get_attribute('href')
            if j.startswith('javascript:'):
                j = j[11:]
            pag.append(j)

    p = -1
    if len(pag) > 0:
        print 'total %s pages found' % (len(pag) + 1)
    urls = []
    statuses = []

    while True:
        lines = driver.find_elements_by_css_selector('tr[class^="SearchResultsGrid"]')
        for line in lines:
            if line.get_attribute('class') == 'SearchResultsGridHeader':
                continue
            tds = line.find_elements_by_css_selector('td')
            if len(tds) < 2:
                continue
            elem = findElement(tds[0], 'a')
            if elem is None:
                continue
            urls.append(elem.get_attribute('href'))
            status = tds[-1].text
            statuses.append(status)
        p += 1
        if p >= len(pag):
            break
        # next page
        print p, len(pag), pag[p]
        driver.execute_script(pag[p])
        waitForLoading(driver)

    print 'total %s items found' % len(urls)

    for i, url in enumerate(urls):
        try:
            data = parseItem(url, writer)
            data.insert(1, statuses[i])
            data.insert(0, fname)
            writerow(writer, data)
        except:
            print traceback.format_exc()

    f.close()

def writerow(writer, row):
    writer.c += 1
    for i, val in enumerate(row):
        writer.write(writer.c, i, val)

driver = None

def main():
    date = datetime.datetime.today()
    delta = datetime.timedelta(days=1)
    date = date + delta
    while date.weekday() != 2:
        date = date + delta
    date = date.strftime('%m/%d/%Y')

    date1 = raw_input("Enter date from (%s)" % date)
    date2 = raw_input("Enter date to (%s)" % date)

    if date1 == '':
        date1 = date
    if date2 == '':
        date2 = date

    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_argument("user-agent=%s" % userAgent)
    #chromeOptions.add_argument("user-data-dir=d:\serv\scrape\chrome")
    global driver
    global shortIDs

    try:
        driver = webdriver.Chrome(chrome_options=chromeOptions)

        shortIDs = False
        parsePage('http://www.bouldercountypt.org/GTSSearch/', 'Boulder', date1, date2) # iframe
        parsePage('http://www.larimer.org/publictrustee/search/index.aspx?ds=1', 'larimer.org', date1, date2)
        parsePage('http://www.wcpto.com/index.aspx', 'wcpto.com', date1, date2)
        parsePage('http://foreclosuresearch.arapahoegov.com/foreclosure/', 'Arapahoe', date1, date2)
        parsePage('http://gts.co.jefferson.co.us/index.aspx', 'Jefferson', date1, date2) # captcha

        shortIDs = True
        parsePage('http://apps.adcogov.org/PTForeclosureSearch/index.aspx?ds=1', 'Adams', date1, date2)

        print "done"
    finally:
        driver.close()

main()
