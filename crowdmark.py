#! python3
# uploads PDF files for Crowdmark assignments
import imapclient, pyzmail, PyPDF2, time, os, bs4
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
pwd = input('Enter your email password: ')
course = input('Enter the math course number (136 or 138): ')  # can change these courses as desired
assert course == '136' or course == '138', 'Wrong course number'
course = 'MATH ' + course
assignment = input('Enter the assignment number: ')
dirname = course + ' A' + assignment
os.chdir('directory')  # go to a directory which will contain the downloaded files 
os.makedirs(dirname, exist_ok=True)
os.chdir(dirname)
client = imapclient.IMAPClient('imap-mail.outlook.com', ssl=True)  # use a different IMAP URL if necessary
client.login('example@email.com', pwd)
client.select_folder('INBOX', readonly=True)
# select the email which contains the scanned pictures and page numbers
UIDs = client.search(['SUBJECT "' + dirname + '" FROM example@email.com'])
assert len(UIDs) == 1, 'email with scanned pages not found'
rawMessages = client.fetch(UIDs, ['BODY[]'])
message = pyzmail.PyzMessage.factory(rawMessages[UIDs[0]][b'BODY[]'])
file = message.mailparts[1]
assert file.type.startswith('application/pdf'), 'attached PDF file not found'
fo = open(dirname+'.pdf', 'wb')
print('Downloading PDF from email...')
fo.write(file.get_payload())
fo.close()
print('Downloaded.')
try:
    text = message.text_part.get_payload().decode(message.text_part.charset)
    numbers = [int(n) for n in text.splitlines()]
except:
    raise Exception('you need to include page numbers in the email message')
fi = open(dirname+'.pdf', 'rb')
reader = PyPDF2.PdfFileReader(fi)
assert sum(numbers) == reader.numPages, 'page numbers don\'t add up'
pageindex = 0
for q in range(len(numbers)):
    n = numbers[q]
    writer = PyPDF2.PdfFileWriter()
    print('Splitting q%s...' % (q+1))
    for i in range(n):
        page = reader.getPage(pageindex)
        page.rotateCounterClockwise(90)
        writer.addPage(page)
        pageindex += 1
    fo = open('q'+str(q+1)+'.pdf', 'wb')
    writer.write(fo)
    fo.close()
fi.close()
# select the email which contains the Crowdmark link
UIDs = client.search(['SUBJECT "FW: ' + course + '" SUBJECT "New assignment: Assignment ' + assignment + '"'])
assert len(UIDs) == 1, 'email with Crowdmark link not found'
rawMessages = client.fetch(UIDs, ['BODY[]'])
message = pyzmail.PyzMessage.factory(rawMessages[UIDs[0]][b'BODY[]'])
html = message.html_part.get_payload().decode(message.html_part.charset)
soup = bs4.BeautifulSoup(html, 'html.parser')
href = soup.select('a')[0].get('href')
client.logout()

# now upload the files to Crowdmark
browser = webdriver.Chrome()
browser.get(href)
WebDriverWait(browser, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'input')))
assert browser.find_elements_by_css_selector('a[class="outline zoom"]') == [], 'files have already been uploaded'
uploadareas = browser.find_elements_by_tag_name('input')
assert len(numbers) == len(uploadareas), 'dont\'t have enough files to upload' 
for i,area in enumerate(uploadareas):
    filename = 'q'+str(i+1)+'.pdf'
    area.send_keys(os.path.abspath(filename))
    time.sleep(0.3)
WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a[class="outline zoom"]')))
browser.find_element_by_css_selector('a[class="outline zoom"]').click()
for j in range(sum(numbers)):
    btn = browser.find_element_by_css_selector('button.icon--refresh')
    for j in range(3):
        btn.click()
        time.sleep(0.2)
    browser.find_element_by_class_name('next-page').click()
browser.find_element_by_css_selector('button.icon--check').click()
btn = browser.find_element_by_css_selector('button[class="button u-push-bottom--2  "]')
btn.click()
print('Assignment submitted.')
# time.sleep(10)
# browser.quit()
