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

te = 0
err = 0


def WriterExcel(worksheet, word, wordtype, sound1, sound2,
                descriptions1, descriptions2, descriptions3,
                examples_text1, examples_text2, examples_text3,
                phons1, phons2, indexS):
    # workbook = xlsxwriter.Workbook('hello.xlsx')
    # worksheet = workbook.add_worksheet()
    indexS = int(indexS)+int(1)
    try:
        worksheet.write('A'+str(indexS), str(word))
        worksheet.write('B'+str(indexS), str(wordtype))
        worksheet.write('C'+str(indexS), str(sound1))
        worksheet.write('D'+str(indexS), str(sound2))
        worksheet.write('E'+str(indexS), str(descriptions1))
        worksheet.write('F'+str(indexS), str(descriptions2))
        worksheet.write('G'+str(indexS), str(descriptions3))
        worksheet.write('H'+str(indexS), str(examples_text1))
        worksheet.write('I'+str(indexS), str(examples_text2))
        worksheet.write('J'+str(indexS), str(examples_text3))
        worksheet.write('K'+str(indexS), str(phons1))
        worksheet.write('L'+str(indexS), str(phons2))
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
        a = 0

        for item in descriptions_ele:
            # descriptions_text.append("mean " + str(a)+": " + item.text)
            # examples_text.append("ex " + str(a)+": " + item.text)
            descriptions_text.append(item.text)
            examples_text.append(item.text)
            if a == 2:
                break

            a = int(a) + int(1)

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
soup = BeautifulSoup(driver.page_source, 'html.parser')
workbook = xlsxwriter.Workbook('hello-2.xlsx')
worksheet = workbook.add_worksheet()

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

            indexStr = WriterExcel(worksheet, word, wordtype, sound[0], sound[1],
                                   descriptions[0], descriptions[1], descriptions[2],
                                   examples[0], examples[1], examples[2],
                                   phons[0], phons[1], te)

            print("----- " + str(indexStr) + " ------")
            # print("-----------")
        except Exception as e:
            err = int(err)+int(1)
            print("Error: "+str(e))
            pass

        te = int(te)+int(1)
        if te == 10:
            workbook.close()
            print("QUIT")
            quit()
# word = driver.find_element_by_xpath(
#     '//*[@id="wordlistsContentPanel"]/ul/li[{}]'.format(i)).get_attribute('data-hw')
# link_word = driver.find_element_by_xpath(
#     '//*[@id="wordlistsContentPanel"]/ul/li[{}]/a'.format(i)).get_attribute('href')


# descriptions, examples, phons = getInforOneWord(link_word)
# print([word, descriptions, examples, phons])
driver.close()
