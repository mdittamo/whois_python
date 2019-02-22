import whois
import time
import xlrd
from colorama import Fore, Back, Style

print(Fore.BLUE)
print"----------------------------------------"
print"  __      __.__             .___        "
print" /  \    /  \  |__   ____   |   | ______"
print" \   \/\/   /  |  \ /  _ \  |   |/  ___/"
print"  \        /|   Y  (  <_> ) |   |\___ \ "
print"   \__/\  / |___|  /\____/  |___/____  >"
print"        \/       \/                  \/ "
print"----------------------------------------"

print (Style.RESET_ALL)

def main():

    while True:
        selection = raw_input ("Enter [s] for single domain, [x] for Excel file or [q] to quit: ")

        if selection =='q':
            print '-----' * 8
            print '------------Exiting program-------------'
            print '-----' * 8
            time.sleep(2.5)
            exit()

        if selection == 's':
            data = raw_input("Enter the domain or [q] to quit: ")
            if data == 'q':
                print '-----' * 8
                print '------------Exiting program-------------'
                print '-----' * 8
                time.sleep(2.5)
                exit()

            else:
                wi = whois.whois(data)
                timestr = time.strftime("%Y%m%d-%H%M%S_")
                name = str(timestr + data)
                f = open('%s.txt' % name, 'w')
                f.write("Results for " + str(data) + "\n")
                f.write(str(wi))
                f.close()
                print '-----' * 8
                print ("Results have been written to " + timestr + data)
                print '-----' * 8
                break

        elif selection == 'x':

           try:
                workbook_name = raw_input('Enter full name of Excel document and extension: ')
                workbook = xlrd.open_workbook(workbook_name)
                worksheet = workbook.sheet_by_index(0)
                columnid = 0;
                for row_index in xrange(worksheet.nrows):
                    entries = worksheet.cell(columnid, 0).value
                    wi = whois.whois(str(entries))
                    timestr = time.strftime("%Y%m%d-%H%M%S_")
                    name = str(timestr + entries)
                    f = open('%s.txt' % name, 'w')
                    f.write("Results for " + str(entries) + "\n")
                    f.write(str(wi))
                    f.close()
                    time.sleep(1)
                    print ("Results have been written to " + timestr + entries)
                    columnid = columnid + 1
                break
           except:
                print('The file does not exist!')
                time.sleep(2.5)
                break
while True:
    main ()
    while True:
        answer = raw_input('Would you like to run again? (y/n): ')
        if answer in ('y', 'n'):
            break
        print 'Invalid input'
    if answer == 'y':
        continue
    else:
        print '-----' * 8
        print '------------Exiting program-------------'
        print '-----' * 8
        time.sleep(2.5)
        break
