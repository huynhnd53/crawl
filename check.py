from selenium import webdriver
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
import xlsxwriter

options = webdriver.ChromeOptions()
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
link_oxford = "https://www.oxfordlearnersdictionaries.com/wordlists/oxford3000-5000?fbclid=IwAR3ItBv_psXmGwngDQAmAG63wkP_YpmX2hRrRnTdlU87Rm06LmJEwaBzVK0"


def getInforOneWord(link_word):
    try:
        driver.get(link_word)
        markup = driver.page_source
        soup = BeautifulSoup(markup, 'html.parser')
        word = soup.findAll('h1', {"class": "headword"})[0].text
        wordtype = soup.findAll('span', {"class": "pos"})[0].text
        phons_br0 = soup.findAll('span', {"class": "phonetics"})[0]
        phons_br = soup.findAll('span', {"class": "phonetics"})[
            0].contents[1].contents[1].text
        sound_phone_br = soup.findAll('span', {"class": "phonetics"})[
            0].contents[1].contents[0]['data-src-mp3']
        phons__n_am = soup.findAll('span', {"class": "phonetics"})[
            0].contents[3].contents[1].text
        sound_phons__n_am = soup.findAll('span', {"class": "phonetics"})[
            0].contents[3].contents[0]['data-src-mp3']
        descriptions_ele = soup.findAll('span', {"class": "def"})
        examples_ele = soup.findAll('ul', {"class": "examples"})
        descriptions_text = []
        examples_text = []
        phons = []
        sound = []
        a = 0
        for item in descriptions_ele:
            descriptions_text.append("mean " + str(a)+": " + item.text)
            examples_text.append("ex " + str(a)+": " + item.text)
            if a == 2:
                break

            a = int(a) + int(1)

        phons.append("phons_br: " + str(phons_br))
        phons.append("phone_n_am: " + str(phons__n_am))
        sound.append("Sound br: "+str(sound_phone_br))
        sound.append("Sound n_Am: "+str(sound_phons__n_am))

        indexTemp = 0

        indexStr = WriterExcel(word, wordtype, sound[0], sound[1],
                               descriptions_text[0], descriptions_text[1], descriptions_text[2],
                               examples_text[0], examples_text[1], examples_text[2],
                               phons[0], phons[1], indexTemp)
        indexTemp = int(indexTemp)+int(1)
        return word, wordtype, sound, descriptions_text, examples_text, phons
    except Exception as e:
        print("Error: "+str(e))
        pass


def WriterExcel(word, wordtype, sound1, sound2,
                descriptions1, descriptions2, descriptions3,
                examples_text1, examples_text2, examples_text3,
                phons1, phons2, index):

    workbook = xlsxwriter.Workbook('hello.xlsx')
    worksheet = workbook.add_worksheet()
    index = int(index)+int(1)
    try:
        worksheet.write('A'+str(index), str(word))
        worksheet.write('B'+str(index), str(wordtype))
        worksheet.write('C'+str(index), str(sound1))
        worksheet.write('D'+str(index), str(sound2))
        worksheet.write('E'+str(index), str(descriptions1))
        worksheet.write('F'+str(index), str(descriptions2))
        worksheet.write('G'+str(index), str(descriptions3))
        worksheet.write('H'+str(index), str(examples_text1))
        worksheet.write('I'+str(index), str(examples_text2))
        worksheet.write('J'+str(index), str(examples_text3))
        worksheet.write('K'+str(index), str(phons1))
        worksheet.write('L'+str(index), str(phons2))
        worksheet.write('M'+str(index), str(index))
        workbook.close()
        return "index"
    except Exception as e:
        print("Error writter: "+str(e))
        pass


driver.get(link_oxford)

te = 0
indexcrr = 0
try:
    href = str(
        "https://www.oxfordlearnersdictionaries.com/definition/english/abuse_1")
    word, wordtype, sound, descriptions, examples, phons = getInforOneWord(
        href)
    print([word, wordtype, sound, descriptions, examples, phons])
    print("Word: ", word)
    print("WordType: ", wordtype)
    print("Sound: ", sound)
    print("Description: ", descriptions)
    print("Example: ", examples)
    print("Phons: ", phons)
    print("index crr: "+str(int(te)+int(1)))
    print("-----------")
except Exception as e:
    print("Error"+str(e))
    pass
# word = driver.find_element_by_xpath(
#     '//*[@id="wordlistsContentPanel"]/ul/li[{}]'.format(i)).get_attribute('data-hw')
# link_word = driver.find_element_by_xpath(
#     '//*[@id="wordlistsContentPanel"]/ul/li[{}]/a'.format(i)).get_attribute('href')


# descriptions, examples, phons = getInforOneWord(link_word)
# print([word, descriptions, examples, phons])
driver.close()
