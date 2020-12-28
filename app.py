from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlsxwriter
import time

class Amazon:
    def __init__(self, email, password):
        self.email = email
        self.password = password
        self.bot = webdriver.Chrome()

    def login(self):
        bot = self.bot
        bot.get("https://www.amazon.com/")
        time.sleep(2)
        
        signInBtn = bot.find_element_by_xpath('//*[@id="nav-link-accountList"]')
        signInBtn.click()
        time.sleep(7)
        
        email = bot.find_element_by_xpath('//*[@id="ap_email"]')
        email.click()
        email.send_keys(self.email)
        time.sleep(5)
        email.send_keys(Keys.RETURN)
        
        password = bot.find_element_by_xpath('//*[@id="ap_password"]')
        password.click()
        password.send_keys(self.password)
        time.sleep(5)
        password.send_keys(Keys.RETURN)
    
    def excel(self, names, prices, dates):
    
        workbook = xlsxwriter.Workbook('Last10Purchases.xlsx')
        worksheet = workbook.add_worksheet()

        worksheet.write("A1", "Name")
        worksheet.write("B1", "Price")
        worksheet.write("C1", "Date")

        for i in range(len(names)):
            worksheet.write(i + 1, 0, names[i])
            worksheet.write(i + 1, 1, float(prices[i]))
            worksheet.write(i + 1, 2, dates[i])

        # change B:11 depending on how much data is in your array
        worksheet.write("E1", "Total")
        worksheet.write_formula("E2", "=SUM(B1:B11)")

        workbook.close()

    def gerOrders(self):
        orderPrice = []
        orderName = []
        orderDate = []
        bot = self.bot
        bot.find_element_by_id('nav-orders').click()
        
        for items in range(3,13):
            orderP = bot.find_element_by_xpath("//*[@id='ordersContainer']/div[" + str(items) + "]/div[2]/div/div[2]/div/div[1]/div/div/div/div[2]/div[4]/span")
            price =  orderP.text
            orderPrice.append(price[1:])

            orderN = bot.find_element_by_xpath("//*[@id='ordersContainer']/div[" + str(items) + "]/div[2]/div/div[2]/div/div[1]/div/div/div/div[2]/div[1]/a")
            name = orderN.text
            orderName.append(name)

            orderD = bot.find_element_by_xpath("//*[@id='ordersContainer']/div[" + str(items) + "]/div[1]/div/div/div/div[1]/div/div[1]/div[2]/span")
            date = orderD.text
            orderDate.append(date)

        # time.sleep(2)

        test.excel(orderName, orderPrice, orderDate)
        time.sleep(10)
        bot.close()

# Enter your email and password
test = Amazon("email", "password")

test.login()
test.gerOrders()



