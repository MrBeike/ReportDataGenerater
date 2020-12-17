# uncompyle6 version 3.7.4
# Python bytecode 3.7 (3394)
# Decompiled from: Python 3.7.3 (v3.7.3:ef4ec6ed12, Mar 25 2019, 21:26:53) [MSC v.1916 32 bit (Intel)]
# Embedded file name: 1029log_test.py
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
print('*****************', '镇江中支调查统计科|0511-85240410', '******************')
print('\n')
dcm_name_lst = {'CLTYJD':'存量同业借贷信息', 
 'TYJDFS':'同业借贷发生额信息', 
 'CLDWDK':'存量单位贷款信息', 
 'DWDKFS':'单位贷款发生额信息', 
 'CLGRDK':'存量个人贷款信息', 
 'GRDKFS':'个人贷款发生额信息', 
 'CLZXDK':'存量专项贷款信息', 
 'DBHTXX':'担保合同信息', 
 'DBWXX':'担保物信息', 
 'CLWTDK':'存量委托贷款信息', 
 'WTDKFS':'委托贷款发生额信息', 
 'JRJGFZ':'金融机构（分支机构）基础信息', 
 'TYKHXX':'同业客户基础信息', 
 'FTYKHX':'非同业单位客户基础信息', 
 'GRKHXX':'个人客户基础信息', 
 'CLTYCK':'存量同业存款信息', 
 'TYCKFS':'同业存款发生额信息', 
 'FTYDWC':'存量非同业单位存款信息', 
 'DWCKFS':'非同业单位存款发生额信息', 
 'CLGRCK':'存量个人存款信息', 
 'GRCKFS':'个人存款发生额信息', 
 'CLZQTZ':'存量债券投资信息', 
 'ZQTZFS':'债券投资发生额信息', 
 'CLZQFX':'存量债券发行信息', 
 'ZQFXFS':'债券发行发生额信息', 
 'CLGQTZ':'存量股权投资信息', 
 'GQTZFS':'股权投资发生额信息', 
 'SPVTZX':'存量特定目的载体投资信息', 
 'SPVFSX':'特定目的载体投资发生额信息'}
dcm_lst = {'CLTYJD':{11:2, 
  12:2,  14:5,  16:5}, 
 'TYJDFS':{13:2, 
  14:2,  16:5,  18:5}, 
 'CLDWDK':{19:2, 
  20:2,  22:5,  24:5}, 
 'DWDKFS':{19:2, 
  20:2,  22:5,  24:5}, 
 'CLGRDK':{14:2, 
  15:2,  17:5,  19:5}, 
 'GRDKFS':{14:2, 
  15:2,  17:5,  19:5}, 
 'CLZXDK':{},  'DBHTXX':{10:2, 
  11:2,  12:2}, 
 'DBWXX':{11:2, 
  13:2,  14:2}, 
 'CLWTDK':{18:2, 
  19:2,  21:5,  22:2}, 
 'WTDKFS':{18:2, 
  19:2,  21:5,  22:2}, 
 'JRJGFZ':{},  'TYKHXX':{15:0, 
  16:0}, 
 'FTYKHX':{8:2, 
  9:2,  10:2,  11:2,  12:0,  23:2,  24:2,  28:0,  29:0}, 
 'GRKHXX':{10:2, 
  11:2,  14:2,  15:2,  19:0,  20:0}, 
 'CLTYCK':{11:2, 
  12:2,  13:5}, 
 'TYCKFS':{11:2, 
  12:2,  15:5}, 
 'FTYDWC':{13:2, 
  14:2,  15:5}, 
 'DWCKFS':{14:2, 
  15:2,  16:5}, 
 'CLGRCK':{12:2, 
  13:2,  14:5}, 
 'GRCKFS':{13:2, 
  14:2,  15:5}, 
 'CLZQTZ':{8:2, 
  9:2,  13:5}, 
 'ZQTZFS':{11:5, 
  20:2,  21:2}, 
 'CLZQFX':{6:0, 
  8:2,  9:2,  10:2,  11:2,  16:5}, 
 'ZQFXFS':{6:0, 
  8:2,  9:2,  10:2,  11:2,  16:5}, 
 'CLGQTZ':{9:2, 
  10:2}, 
 'GQTZFS':{10:2, 
  11:2}, 
 'SPVTZX':{12:2, 
  13:2}, 
 'SPVFSX':{13:2, 
  14:2}, 
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

def xlsx2dat(file_name: str, file_lst: list):
    sep = '|'
    tot_row = 0
    num = 0
    txt_file = file_name + '.dat'
    bw_type = file_name.split('_')[1]
    if dcm_lst.get(bw_type, None) == None:
        print('本程序仅支持部分数据报文的生产操作,文件列表如下：')
        print(dcm_name_lst)
        print('--------------------------------------------------------------------------------')
        print(file_name, '暂不支持报文生产操作！')
        print('请核实报表对应字符串（如字母大小写）是否填报错误，修改后重新执行本程序！')
        os.system('pause')
        os._exit(0)
    elif len(file_lst) == 1:
        filename = file_lst[0]
        workbook = xlrd.open_workbook(filename)
        sheet1 = workbook.sheet_by_index(0)
        nrows = sheet1.nrows
        ncols = sheet1.ncols
        if nrows == 1:
            open(txt_file, 'w', encoding='utf-8')
            print('Excel文件中共有记录', nrows, '条；其中数据记录', nrows - 1, '条。')
        else:
            print('累计读取1个文件', 'Excel文件中共有记录', nrows, '条；其中数据记录', nrows - 1, '条。')
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

    else:
        with open(txt_file, 'w', encoding='utf-8') as (txtfile):
            for filename in file_lst:
                num += 1
                workbook = xlrd.open_workbook(filename)
                sheet1 = workbook.sheet_by_index(0)
                nrows = sheet1.nrows
                ncols = sheet1.ncols
                tot_row += nrows
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

                if num < len(file_lst):
                    txtfile.write('\n')

            print('累计读取', len(file_lst), '个文件', 'Excel文件中共有记录', tot_row, '条；其中数据记录', tot_row - len(file_lst), '条。')


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
    if count == 0:
        md5 = ''
    else:
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


# filelst = os.listdir()
# for each in filelst:
#     if not each[-4:] == '.dat':
#         if each[-4:] == '.log':
#             pass
#         os.remove(each)

print('--------------------', '（2）程序根据xlsx文件，准备生成dat文件', '--------------------')
xlsx_lst = {}
sep_name = '_'
filelst = os.listdir()
for each in filelst:
    if each[-5:] == '.xlsx' and each.find('_') != -1:
        num += 1
        if len(each.split('_')) == 4:
            dat_name = sep_name.join(each.split('_')[:3])
            if dat_name in xlsx_lst.keys():
                xlsx_lst[dat_name].append(each)
            else:
                xlsx_lst[dat_name] = [
                 each]
        else:
            dat_name = each.split('.')[0]
            xlsx_lst[dat_name] = [each]

jishu = 0
if num == 0:
    print('本程序所在文件夹下不存在格式为xlsx的文件，请放入xlsx文件后重新运行本程序！')
else:
    for each in list(xlsx_lst.keys()):
        jishu += 1
        print(str(jishu), '.正在生成', each, '文件的dat文件')
        xlsx2dat(each, xlsx_lst[each])

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
# for each in filelst:
#     if not each[-4:] == '.dat':
#         if each[-4:] == '.log':
#             pass
#         os.remove(each)

os.system('pause')
# okay decompiling 1029log_test.pyc
