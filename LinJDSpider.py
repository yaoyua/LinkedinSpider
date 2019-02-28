#Written by Frank Lin in 2017.
from selenium import webdriver
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import os
from fileFunc import *
from SplitJD import *
import xlsxwriter
import sys


def main():

    # username = input("Linkedin Username: ")
    # password = input("Linkedin Password: ")
    # path = input("input file: ")
    # fileName = input("output file: ")

    if len(sys.argv[1:]) < 4:
        print("Need more arguments!")
        exit(-1)
    elif len(sys.argv[1:]) > 4:
        print("Too many arguments!")
        exit(-1)
    else:
        username = sys.argv[1]
        password = sys.argv[2]
        path = sys.argv[3]
        fileName = sys.argv[4]

    if os.path.isfile(fileName):
        print('old file')
    else:
        print('new file')

    workbook = xlsxwriter.Workbook(fileName)

    chrome_options = webdriver.ChromeOptions()
    driver = webdriver.Chrome('chromedriver', chrome_options=chrome_options)
    #driver = webdriver.Firefox()


    driver.get('https://www.linkedin.com/')
    driver.set_page_load_timeout(100)
    driver.find_element_by_name("session_key").send_keys(username)
    driver.find_element_by_name("session_password").send_keys(password)
    driver.find_element_by_id("login-submit").click()


    search_List = open_file(path)
    for search_pro in search_List:

        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, 'Job Title')
        worksheet.write(0, 1, 'Company Name')
        worksheet.write(0, 2, 'Company Link')
        worksheet.write(0, 3, 'Job Description')
        worksheet.write(0, 4, 'Responsibilites')
        worksheet.write(0, 5, 'Qualifications')
        
        search_id = search_pro[0]
        search_job_title = search_pro[1]
        search_location = search_pro[2]
        # search_first_name = search_pro[1]
        # search_last_name = search_pro[2]
        keyWordtitle = search_job_title.replace(" ", "%20")
        keyWordloc = search_location.replace(" ", "%20")

        searchURL = 'https://www.linkedin.com/jobs/search/?keywords=' + keyWordtitle

        if len(keyWordloc) != 0:
            searchURL = searchURL + '&location=' + keyWordloc

        print("URL = ", searchURL)

        links = []
        count = 0

        while True:
            pageURL = searchURL + '&start=' + str(count)
            print("pageURL = ", pageURL)
            driver.get(pageURL)

            #test whether result exists
            if driver.find_elements_by_css_selector(".search-no-results"):
                print("no results")
                break

            results = driver.find_elements_by_xpath("//div[@data-control-name='A_jobssearch_job_result_click']")

            print(len(results))
            if len(results) == 0:
                break

            for r in results:
                link = r.get_attribute("data-job-id")
                print(count, link)
                count += 1
                links.append(link)

                sleep(1)

                # print("text = ", text)
        myCount = 1
        for link in links:
            number_id = link[32:]
            print("number = ", myCount)
            url = 'https://www.linkedin.com/jobs/search/?currentJobId=' + number_id + '&keywords=' + keyWordtitle
            if len(keyWordloc) != 0:
                url += '&location=' + keyWordloc
            print("url = ", url)
            driver.get(url)

            sleep(2)

            # Job Title
            title = driver.find_element_by_xpath(
                "//h1[@class='jobs-details-top-card__job-title t-20 t-black t-normal']").text
            worksheet.write(myCount, 0, title)

            # Company Name & Company Link
            try:
                company = driver.find_element_by_xpath("//a[@class='jobs-details-top-card__company-url ember-view']")
                company_name = company.text
                company_link = company.get_attribute('href')
                worksheet.write(myCount, 1, company_name)
                worksheet.write(myCount, 2, company_link)
            except:
                company = driver.find_element_by_xpath(
                    "//h3[@class='jobs-details-top-card__company-info t-14 t-black--light t-normal mt1']")
                company_name = company.text
                worksheet.write(myCount, 1, company_name)

            # Job Description
            elements = driver.find_element_by_xpath(
                "//div[@class='jobs-box__html-content jobs-description-content__text t-14 t-black--light t-normal']")
            element = elements.find_element_by_tag_name('span')
            text = element.text

            # FIXME
            if len(text) <= 5:
                print("Nothing's url = ", url)

            resp, quali = Split(text)
            worksheet.write(myCount, 3, text)
            worksheet.write(myCount, 4, resp)
            worksheet.write(myCount, 5, quali)

            myCount += 1

            sleep(2)

        workbook.close()

if __name__ == '__main__':
    main()
