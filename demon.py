import os
import sys
import xlrd
import xlwt
import getpass
import signal
import win32api
import win32print

options = ["Separately print", "Continuously print"]
printers = []
row_num = 0

dict_num = {"0":u"零","1":u"壹","2":u"贰","3":u"叁","4":u"肆","5":u"伍","6":u"陆","7":u"柒","8":u"扒","9":u"玖"}
dict_unit = {0:u"圆",1:u"拾",2:u"佰",3:u"仟",4:u"万"}
space_num = [80, 10, 20, 10, 5, 41, 5, 20, 5, 50, 85]

default_printer = str(win32print.GetDefaultPrinter())
invoice_printer = ""

def signal_handler(signum, frame):
    if signum == signal.SIGINT:
        print("Keyboard Interrupt")
    else:
        print("Recv signal %d"%(signal))

    win32print.SetDefaultPrinter(default_printer)
    sys.exit(0)

def get_printers():
    t = win32print.EnumPrinters(win32print.PRINTER_ENUM_NAME, None, 4)
    for i in range(len(t)):
        printers.append(t[i]["pPrinterName"])

def prt_option(options):
    print("Options list:")
    for i in range(len(options)):
        print("%d. %s"%(i,options[i]))

    option = input("Please input the number of the option:")

    return option

def init_env():
    global row_num
    global invoice_printer

    prt_type = int(prt_option(options))
    if prt_type not in range(2):
        print("Invalid option: ", prt_type)
        return False

    f_path = input("Please input full file path: ")
    if (not os.path.exists(f_path)) or not (os.path.isfile(f_path)):
        print("File not exist: ", f_path)
        return False

    if (prt_type == 0):
        row_num = int(input("Please input row num:"))
        if row_num <= 0:
            print("Invalid row num:", row_num)
            return False

    option = int(prt_option(printers))
    invoice_printer = printers[option]
    if invoice_printer is "":
        print("Invalid invoice printer:", invoice_printer)
        return False
    
    print("You have selected: \n\t", options[prt_type]) 
    print("\t",f_path)
    if (prt_type == 0):
        print("\t row num:", row_num)
    print("\t printer:", invoice_printer)
 
    return f_path

def exchange_num(element):
    num = []
    s = ''

    lennum = len(element) - 1
    for i in element:
        num.append(dict_num[i])
        num.append(dict_unit[lennum])
        lennum -= 1
    num.append(u"整")
    s = u'合计人民币(大写)：' + ''.join(num)
    
    return s

def reorg_data(s_list):
    s_list = s_list[:6] + s_list[13:]
    element = s_list[6]
    del s_list[6]
    s_list.append(element)

    s_list[1] = str(int(s_list[1]))
    s_list[2] = s_list[2].replace("-", "  ")
    element = list(str(int(s_list[5])))
    elem1 = exchange_num(element)
    elem2 = u"￥:" + str(s_list[5])
    s_list.insert(6, elem1)
    s_list.insert(7, elem2)
    s_list[5] = str(s_list[5])

    return s_list

def gen_space(num):
    s = ''

    while num:
        s = s + " "
        num = num - 1
    
    return s

def draw_inv(d_list):
    for i in range(len(d_list)):
        d_list[i] = gen_space(space_num[i]) + d_list[i]

    d_list[0] = '\n\n\n' + d_list[0] + '\n'
    d_list[2] = d_list[2] + '\n'
    d_list[1] = d_list[1] + d_list[2]
    d_list[3] = d_list[3] + '\n\n\n'
    d_list[5] = d_list[5] + '\n'
    d_list[4] = d_list[4] + d_list[5]
    d_list[7] = d_list[7] + '\n'
    d_list[6] = d_list[6] + d_list[7]
    d_list[9] = d_list[9] + '\n\n\n'
    d_list[8] = d_list[8] + d_list[9]

    del d_list[9]
    del d_list[7]
    del d_list[5]
    del d_list[2]

    return d_list

def get_data(fd_sheet, row):
    s_list = fd_sheet.row_values(row)
    if not (u"普通高校" in s_list[4]):
        return []
    
    d_list = reorg_data(s_list)
    d_list = draw_inv(d_list)

    return d_list

def continu_print(fd_sheet):
    for i in range(len(fd_sheet.row_values(6))):
        print(str(fd_sheet.row_values(6)[i]))

def create_tmpsheet(path):
    x = xlwt.Workbook(encoding='utf-8')
    sheet = x.add_sheet('tmp')

    sheet.col(0).width = 256*120
    height = xlwt.easyxf('font:height 3600')
    sheet.row(0).set_style(height)

    #VERT_TOP       = 0x00
    #VERT_CENTER    = 0x01    
    #VERT_BOTTOM    = 0x02    
    #HORZ_LEFT      = 0x01    
    #HORZ_CENTER    = 0x02    
    #HORZ_RIGHT     = 0x03    
    alignment = xlwt.Alignment()
    alignment.horz = 0x01
    alignment.vert = 0x00
    alignment.wrap = 1

    font = xlwt.Font()
    font.name = u"宋体"
    font.height = 240

    style = xlwt.XFStyle()
    style.font = font
    style.alignment = alignment
        
    s = u"".join(raw_data)
    sheet.write(0, 0, s, style)

    x.save(path)

def print_windows(data, inv_printer):
    win32print.SetDefaultPrinter(inv_printer)
    win32api.ShellExecute(0, "print", data, None, ".", 0)

if __name__ == '__main__':
    #Ctrl + C to stop
    signal.signal(signal.SIGINT, signal_handler)
    get_printers()

    #Get source file and printer
    f_path = init_env()
    if not f_path:
        sys.exit(0)

    chk = input("Are you sure?(Y/n):")
    if (chk == "n") or (chk == "N"):
        sys.exit(0)

    x1 = xlrd.open_workbook(f_path)
    sheet1 = x1.sheet_by_index(0)

    path = "C:\\Users\\" + getpass.getuser() + "\\Desktop\\tmp.xls"
    if row_num > 0:
        raw_data = get_data(sheet1, row_num)
        if raw_data:
            create_tmpsheet(path)
            print_windows(path, invoice_printer)
    else:
        for i in range(sheet1.nrows):
            raw_data = get_data(sheet1, i)
            if raw_data:
                create_tmpsheet(path)
                print_windows(path, invoice_printer)


    print("The end")
