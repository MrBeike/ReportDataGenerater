# uncompyle6 version 3.7.4
# Python bytecode 3.7 (3394)
# Decompiled from: Python 3.7.3 (v3.7.3:ef4ec6ed12, Mar 25 2019, 21:26:53) [MSC v.1916 32 bit (Intel)]
# Embedded file name: 0928log_test.py
import hashlib, os, time, zipfile, xlrd
from datetime import date
sep = '|'
num = 0
txt_bmlst = {}
print('\n')
print('******************************', '小程序介绍', '******************************')
print('一、程序用途')
print('本程序用于对金融机构报送的Excel（xlsx）报文文件生成dat、log文件，并对dat文件、log文件进行压缩。')
print('---------------------------------------------------------------------------')
print('二、操作步骤')
print('第一步：新建一个文件夹。')
print('第二步：将本程序与需要生成报文文件的Excel(xlsx)文件格式（支持多个文件）放在该文件中；')
print('第三步：双击本程序。')
print('******************', '中国人民银行南京分行调查统计处', '*******************')
print('*****************', '镇江中支调查统计科', '******************')
print('\n')
dcm_lst = {'CLDWDK':{19:2, 
  20:2,  22:5,  24:5}, 
 'DWDKFS':{19:2, 
  20:2,  22:5,  24:5}, 
 'DBHTXX':{10:2, 
  11:2,  12:2}, 
 'DBWXX':{11:2, 
  13:2,  14:2}, 
 'JRJGFZ':{},  'FTYKHX':{8:2, 
  9:2,  10:2,  11:2,  12:0,  23:2,  24:2,  28:0,  29:0}, 
 'WQYBJY':{7: 2}}

def xjml(wjj_name):
    filepath = os.getcwd() + '\\' + wjj_name
    try:
        os.mkdir(wjj_name)
    except:
        filelst = os.listdir(wjj_name)
        if len(filelst) != 0:
            for eachfile in filelst:
                path_file = os.path.join(filepath, eachfile)
                os.remove(path_file)

    print('--------', '（1）程序新建了“', wjj_name, '”用于存放ZIP格式的压缩文件！', '--------')
    print('\n')


xjml('ZIP压缩文件夹')

def xlsx2dat(filename):
    sep = '|'
    workbook = xlrd.open_workbook(filename)
    sheet1 = workbook.sheet_by_index(0)
    nrows = sheet1.nrows
    ncols = sheet1.ncols
    txt_file = filename.split('.')[0] + '.dat'
    bw_type = filename.split('_')[1]
    if dcm_lst.get(bw_type, None) == None:
        print('本程序仅支持部分数据报文的生产操作,文件列表如下：')
        print(list(dcm_lst.keys()))
        print(filename, '暂不支持报文生产操作！')
        print('请核实报表对应字符串（如字母大小写）是否填报错误，修改后重新执行本程序！')
        os.system('pause')
        os._exit(0)
    elif nrows == 1:
        open(txt_file, 'w', encoding='utf-8')
        print('Excel文件中共有记录', nrows, '条；其中数据记录', nrows - 1, '条。')
    else:
        print('Excel文件中共有记录', nrows, '条；其中数据记录', nrows - 1, '条。')
        with open(txt_file, 'w', encoding='utf-8') as (txtfile):
            for row_idx in range(1, nrows):
                contents = []
                for col_idx in range(ncols):
                    cell_value = sheet1.cell(row_idx, col_idx).value
                    if sheet1.cell(row_idx, col_idx).ctype == 3:
                        data_value = xlrd.xldate_as_tuple(cell_value, 0)
                        tmp = date(*data_value[:3]).strftime('%Y-%m-%d')
                        contents.append(str(tmp))
                    elif sheet1.cell(row_idx, col_idx).ctype == 1:
                        contents.append(str(cell_value))
                    elif sheet1.cell(row_idx, col_idx).ctype == 2:
                        if col_idx + 1 in list(dcm_lst[bw_type].keys()):
                            ret_val = dcm_lst[bw_type].get(col_idx + 1)
                            if ret_val == 2:
                                tmp = '{:.2f}'.format(cell_value)
                            else:
                                if ret_val == 5:
                                    tmp = '{:.5f}'.format(cell_value)
                                else:
                                    if ret_val == 0:
                                        tmp = int(cell_value)
                                    contents.append(str(tmp))
                        else:
                            contents.append(str(int(cell_value)))
                    else:
                        contents.append(str(cell_value))

                txtfile.write(sep.join(contents))
                if row_idx <= nrows - 2:
                    txtfile.write('\n')


def getfilemd5(filename):
    time1 = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    org_file = open(os.getcwd() + os.sep + filename, 'rb')
    count = len(org_file.readlines())
    myhash = hashlib.md5()
    b = open(os.getcwd() + os.sep + filename, 'rb').read()
    myhash.update(b)
    org_file.close()
    size = os.path.getsize(filename)
    file = filename.split('\\')[(-1)]
    md5 = myhash.hexdigest()
    optname = filename.split('.')[0] + '.log'
    with open((os.getcwd() + os.sep + optname), 'w', encoding='utf8') as (file):
        file.write(filename)
        file.write('\n')
        file.write(md5)
        file.write('\n')
        file.write(str(size))
        file.write('\n')
        file.write(time1)
        file.write('\n')
        file.write(str(count))


filelst = os.listdir()
for each in filelst:
    if not each[-4:] == '.dat':
        if each[-4:] == '.log':
            pass
        os.remove(each)

print('--------------------', '（2）程序根据xlsx文件，准备生成dat文件', '--------------------')
filelst = os.listdir()
for each in filelst:
    if each[-5:] == '.xlsx':
        num += 1

jishu = 0
if num == 0:
    print('本程序所在文件夹下不存在格式为xlsx的文件，请放入xlsx文件后重新运行本程序！')
else:
    for each in filelst:
        if each[-5:] == '.xlsx':
            jishu += 1
            print(str(jishu), '.正在生成', each, '文件的dat文件')
            xlsx2dat(each)

    print('dat文件生成完毕！')
print('\n')
print('--------------------', '（3）程序根据dat文件，准备生成log文件', '--------------------')
filelst = os.listdir()
for each in filelst:
    if each[-4:] == '.dat':
        num += 1

jishu = 0
if num == 0:
    print('本程序所在文件夹下不存在格式为dat的文件，请放入dat文件后重新运行本程序！')
else:
    for each in filelst:
        if each[-4:] == '.dat':
            jishu += 1
            print(str(jishu), '.正在生成', each, '文件的log文件')
            getfilemd5(each)

    print('log文件生成完毕！')
print('\n')
print('--------------', '（4）程序根据生成的dat、log文件，准备生成压缩文件', '--------------')

def writeAllFileToZip():
    jishu = 0
    filelst = []
    absDir = os.getcwd()
    for file in os.listdir(absDir):
        if file[-4:] == '.dat':
            filelst.append(file[:len(file) - 4])

    for f in filelst:
        jishu += 1
        print(str(jishu), '.正在生成', f, '压缩文件')
        zipFile = zipfile.ZipFile(os.getcwd() + os.sep + 'ZIP压缩文件夹' + os.sep + f + '.zip', 'w', zipfile.ZIP_DEFLATED)
        f_txt = f + '.dat'
        f_log = f + '.log'
        try:
            zipFile.write((absDir + os.sep + f + '.dat'), arcname=f_txt)
            zipFile.write((absDir + os.sep + f + '.log'), arcname=f_log)
            zipFile.close()
        except:
            print('在执行文件压缩操作时，dat文件或log文件丢失，请至“数据处理文件夹”查看具体情况！')

    print('搞定啦，累计压缩：', len(filelst), '个文件。请至“ZIP压缩文件夹”中查看！')


writeAllFileToZip()
filelst = os.listdir()
for each in filelst:
    if not each[-4:] == '.dat':
        if each[-4:] == '.log':
            pass
        os.remove(each)

os.system('pause')
# okay decompiling 0928log_test.pyc
