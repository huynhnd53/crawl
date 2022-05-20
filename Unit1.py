from selenium import webdriver
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
import xlsxwriter
# 1. declare web driver
options = webdriver.ChromeOptions()
# options.add_experimental_option('excludeSwitches', ['enable-logging'])
# options.add_argument("start-maximized")
# options.add_argument("disable-infobars")
# options.add_argument("--disable-extensions")
# options.add_argument("--disable-blink-features=AutomationControlled")
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
# driver=webdriver.Chrome()
link_oxford = "https://www.oxfordlearnersdictionaries.com/wordlists/oxford3000-5000"
# tag=[]
# driver.get(link_oxford)
# words=[]
# link_descriptions=[]
# for i in range(1,3000):
#     word=driver.find_element_by_xpath('//*[@id="wordlistsContentPanel"]/ul/li[{}]'.format(i)).get_attribute('data-hw')
#     link=driver.find_element_by_xpath('//*[@id="wordlistsContentPanel"]/ul/li[{}]/a'.format(i)).get_attribute('href')
#     words.append(word)
#     link_descriptions.append(link)
#     if i==5:
#         print(words)
#         print(link_descriptions)
# for link in link_descriptions:
# i (1-3001) (3000 words)




def WriterExcel(worksheet, word, wordtype, sounds, descriptions, examples_text, phons, indexS):
    # workbook = xlsxwriter.Workbook('hello.xlsx')
    # worksheet = workbook.add_worksheet()
    


    indexS = int(indexS)+int(1)
    try:
        worksheet.write('A'+str(indexS), str(word)) #for WORD
        worksheet.write('B'+str(indexS), str(wordtype)) #for WORDTYPE
        
        xS = 1
        for itemS in sounds:
            if xS == 1:
                worksheet.write('D'+str(indexS), str(itemS)) #for SOUND EN-UK
            elif xS == 2:
                worksheet.write('F'+str(indexS), str(itemS)) #for SOUND EN-US
            print("sounds" + str(xS) + ": " + itemS)
            xS = int(xS) + int(1)

        xP = 1
        for itemP in phons:
            if xP == 1:
                worksheet.write('C'+str(indexS), str(itemP)) #for PHONS EN-UK
            elif xP == 2:
                worksheet.write('E'+str(indexS), str(itemP)) #for PHONS EN-US
            print("phons" + str(xP) + ": " + itemP)
            xP = int(xP) + int(1)

        xD = 1
        for itemD in descriptions:
            if xD == 1:
                worksheet.write('G'+str(indexS), str(itemD)) #for DESCRIPTION 1
            elif xD == 2:
                worksheet.write('H'+str(indexS), str(itemD)) #for DESCRIPTION 2
            elif xD == 3:
                worksheet.write('I'+str(indexS), str(itemD)) #for DESCRIPTION 3
            print("descriptions" + str(xD) + ": " + itemD)
            xD = int(xD) + int(1)

        xEx = 1
        for itemEx in examples_text:
            if xEx == 1:
                worksheet.write('J'+str(indexS), str(itemEx)) #for EXAMPLE 1
            elif xEx == 2:
                worksheet.write('K'+str(indexS), str(itemEx)) #for EXAMPLE 2
            elif xEx == 3:
                worksheet.write('L'+str(indexS), str(itemEx)) #for EXAMPLE 3
            print("examples_text" + str(xEx) + ": " + itemEx)
            xEx = int(xEx) + int(1)

        worksheet.write('M'+str(indexS), str(indexS))
        # workbook.close()
        return indexS
    except Exception as es:
        print("Error WriterExcel: "+str(es))
        pass


def getInforOneWord(link_word):

    try:
        driver.get(link_word)
        markup = driver.page_source
        soup = BeautifulSoup(markup, 'html.parser')
        word = soup.findAll('h1', {"class": "headword"})[0].text
        wordtype = soup.findAll('span', {"class": "pos"})[0].text
        phons_br0 = soup.findAll('span', {"class": "phonetics"})[0]
        phons_br = soup.findAll('span', {"class": "phonetics"})[0].contents[1].contents[1].text
        sound_phone_br = soup.findAll('span', {"class": "phonetics"})[0].contents[1].contents[0]['data-src-mp3']
        phons__n_am = soup.findAll('span', {"class": "phonetics"})[0].contents[3].contents[1].text
        sound_phons__n_am = soup.findAll('span', {"class": "phonetics"})[0].contents[3].contents[0]['data-src-mp3']
        descriptions_ele = soup.findAll('span', {"class": "def"})
        examples_ele = soup.findAll('ul', {"class": "examples"})
        descriptions_text = []
        examples_text = []
        phons = []
        sound = []

        countIdxDes = 0
        for item in descriptions_ele:
            descriptions_text.append(item.text)
            if countIdxDes == 2:
                break
            countIdxDes = int(countIdxDes) + int(1)

        countIdxEx = 0
        for item in examples_ele:
            examples_text.append(item.text)
            if countIdxEx == 2:
                break
            countIdxEx = int(countIdxEx) + int(1)

        # phons.append("phons_br: " + str(phons_br))
        # phons.append("phone_n_am: " + str(phons__n_am))
        # sound.append("Sound br: "+str(sound_phone_br))
        # sound.append("Sound n_Am: "+str(sound_phons__n_am))

        phons.append(str(phons_br))
        phons.append(str(phons__n_am))
        sound.append(str(sound_phone_br))
        sound.append(str(sound_phons__n_am))

        return word, wordtype, sound, descriptions_text, examples_text, phons
    except Exception as ex:
        err = int(err)+int(1)
        print("Error getInforOneWord: "+str(ex))
        pass


driver.get(link_oxford)
te = 1
err = 0

soup = BeautifulSoup(driver.page_source, 'html.parser')
workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()
# Add title of collumn                   
worksheet.write('A1', str("WORD"))
worksheet.write('B1', str("WORDTYPE"))
worksheet.write('C1', str("PHONS EN-UK"))
worksheet.write('D1', str("SOUND EN-UK"))
worksheet.write('E1', str("PHONS EN-US"))
worksheet.write('F1', str("SOUND EN-US"))
worksheet.write('G1', str("DESCRIPTION 1"))
worksheet.write('H1', str("DESCRIPTION 2"))
worksheet.write('I1', str("DESCRIPTION 3"))
worksheet.write('J1', str("EXAMPLE 1"))
worksheet.write('K1', str("EXAMPLE 2"))
worksheet.write('L1', str("EXAMPLE 3"))
worksheet.write('M1', str("Index"))

for ultag in soup.find_all('ul', {'class': 'top-g'}):
    for litag in ultag.find_all('li'):
        try:
            href = str("https://www.oxfordlearnersdictionaries.com/") + litag.contents[1]['href']
            word, wordtype, sound, descriptions, examples, phons = getInforOneWord(href)

            # print([word, wordtype, sound, descriptions, examples, phons])
            print("Word: ", word)
            print("WordType: ", wordtype)
            print("Sound: ", sound)
            print("Description: ", descriptions)
            print("Example: ", examples)
            print("Phons: ", phons)
            print("index crr: "+str(int(te)+int(1)))
            print("Num Error: "+str(err))

            indexStr = WriterExcel(worksheet, word, wordtype, sound, descriptions ,examples, phons, te)
            print("----- " + str(indexStr) + " ------")
            # print("-----------")
        except Exception as e:
            err = int(err)+int(1)
            print("Error: "+str(e))
            pass

        te = int(te)+int(1)
        
        #crawl number of word
        if te == 25:
            workbook.close()
            print("QUIT")
            quit()

#Crawl all         
#workbook.close()


# word = driver.find_element_by_xpath(
#     '//*[@id="wordlistsContentPanel"]/ul/li[{}]'.format(i)).get_attribute('data-hw')
# link_word = driver.find_element_by_xpath(
#     '//*[@id="wordlistsContentPanel"]/ul/li[{}]/a'.format(i)).get_attribute('href')


# descriptions, examples, phons = getInforOneWord(link_word)
# print([word, descriptions, examples, phons])
driver.close()
