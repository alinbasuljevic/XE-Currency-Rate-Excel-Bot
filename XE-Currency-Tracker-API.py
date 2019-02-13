import requests, json, time, xlwt, pyfiglet,termcolor, colorama
from termcolor import cprint
from pyfiglet import figlet_format
from bs4 import BeautifulSoup as bs
colorama.init()

def general_request():
    start_currency = input('Enter Starting Currency Abbreviation: ')
    s = requests.Session()
    product_url = s.get('https://www.xe.com/currencytables/?from=USD')
    soup = bs(product_url.text, 'lxml')
    list_of_currencies = []
    for data in soup.find_all('a'):
        list_of_currencies.append(data.text)

    del list_of_currencies[0:4]
    del list_of_currencies[167:193]

    print('Total Number of Currency Exchange Rates: [{}]'.format(len(list_of_currencies)))

    final_num_of_tasks = []
    for i in range(len(list_of_currencies)):
        exchange_task = [start_currency, list_of_currencies[i]]
        if exchange_task in final_num_of_tasks:
            pass
        else:
            final_num_of_tasks.append(exchange_task)
    
    book = xlwt.Workbook()
    SHEET = book.add_sheet('XE Currency Rates', cell_overwrite_ok=True)
    SHEET.write(0,0, 'Currency Exchanges')
    SHEET.write(0,1, 'Average 30-Day Exchange Rates')
    SHEET.write(0,2, 'Average 90-Day Exchange Rates')

    for i in range(len(final_num_of_tasks)):
        api_url = 'https://www.xe.com/api/stats.php?fromCurrency={}&toCurrency={}'.format(final_num_of_tasks[i][0], final_num_of_tasks[i][1])
        s = requests.Session()
        request = s.get(api_url)
        json_data = json.loads(request.text)
        print ('Average 30-Day Exchange Rate [{}-{}]: [{}]'.format(final_num_of_tasks[i][0], final_num_of_tasks[i][1], json_data['payload']['Last_30_Days']['average']))
        print ('Average 90-Day Exchange Rate [{}-{}]: [{}]'.format(final_num_of_tasks[i][0], final_num_of_tasks[i][1], json_data['payload']['Last_90_Days']['average']))
        SHEET.write(i+1, 0, '[{}-{}]'.format(final_num_of_tasks[i][0], final_num_of_tasks[i][1]))
        SHEET.write(i+1, 1, json_data['payload']['Last_30_Days']['average'])
        SHEET.write(i+1, 2, json_data['payload']['Last_30_Days']['average'])
        i += 1

    book.save('XE-Currency-Tracker.xls')
    print ('Saved to Excel Sheet!')


    


def specific_requests():
    num_of_currencies = int(input('How many rates would you like to fetch?: '))
    final_num_of_tasks = []

    for i in range(num_of_currencies):
        start_currency = input('Enter Starting Currency Abbreviation [{}/{}]: '.format(i+1,num_of_currencies))
        ending_currency = input('Enter Ending Currency Abbreviation [{}/{}]: '.format(i+1, num_of_currencies))
        exchange_task = [start_currency, ending_currency]
        if exchange_task in final_num_of_tasks:
            pass
        else:
            final_num_of_tasks.append(exchange_task)

    book = xlwt.Workbook()
    SHEET = book.add_sheet('XE Parser Currency Rates', cell_overwrite_ok=True)
    SHEET.write(0,0, 'Currency Exchanges')
    SHEET.write(0,1, 'Average 30-Day Exchange Rates')
    SHEET.write(0,2, 'Average 90-Day Exchange Rates')

    for i in range(len(final_num_of_tasks)):
        api_url = 'https://www.xe.com/api/stats.php?fromCurrency={}&toCurrency={}'.format(final_num_of_tasks[i][0], final_num_of_tasks[i][1])
        s = requests.Session()
        request = s.get(api_url)
        json_data = json.loads(request.text)
        print ('Average 30-Day Exchange Rate [{}-{}]: [{}]'.format(final_num_of_tasks[i][0], final_num_of_tasks[i][1], json_data['payload']['Last_30_Days']['average']))
        print ('Average 90-Day Exchange Rate [{}-{}]: [{}]'.format(final_num_of_tasks[i][0], final_num_of_tasks[i][1], json_data['payload']['Last_90_Days']['average']))
        SHEET.write(i+1, 0, '[{}-{}]'.format(final_num_of_tasks[i][0], final_num_of_tasks[i][1]))
        SHEET.write(i+1, 1, json_data['payload']['Last_30_Days']['average'])
        SHEET.write(i+1, 2, json_data['payload']['Last_30_Days']['average'])
        i += 1

    book.save('XE-Currency-Tracker.xls')
    print ('Saved to Excel Sheet!')


if __name__ == "__main__":
    cprint(figlet_format('XE Currency Rate Excel Bot V2 ', font='speed'), 'cyan', attrs=['bold'])
    print ('Enter [1] for General Excel Sheet Containing All Possible Conversion Rates')
    print ('Enter [2] for Specific Excel Sheet Containing Defined Conversion Rates')
    menu_selector = int(input('Enter Menu Option: '))
    if menu_selector == 1:
        general_request()
    elif menu_selector == 2:
        specific_request()
    else:
        pass

    