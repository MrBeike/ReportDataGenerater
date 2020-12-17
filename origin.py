import hashlib
import os
import time
import zipfile
import xlrd
from datetime import date

# 基本参数定义
sep = '|'
num = 0
txt_bmlst = {}
wjj_name = 'ZIP压缩文件夹'
# 获取当前工作路径
workDir = os.path.abspath(os.path.dirname(__file__))

# 数据表名称
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

# 数据精度定义,格式为'报表代码':{列数:精度}
dcm_lst = {'CLTYJD': {11: 2,12: 2,  14: 5,  16: 5},
           'TYJDFS': {13: 2,14: 2,  16: 5,  18: 5},
           'CLDWDK': {19: 2,20: 2,  22: 5,  24: 5},
           'DWDKFS': {19: 2,20: 2,  22: 5,  24: 5},
           'CLGRDK': {14: 2,15: 2,  17: 5,  19: 5},
           'GRDKFS': {14: 2,15: 2,  17: 5,  19: 5},
           'CLZXDK': {},  
           'DBHTXX': {10: 2,11: 2,  12: 2},
           'DBWXX': {11: 2,13: 2,  14: 2},
           'CLWTDK': {18: 2,19: 2,  21: 5,  22: 2},
           'WTDKFS': {18: 2,19: 2,  21: 5,  22: 2},
           'JRJGFZ': {},  
           'TYKHXX': {15: 0, 16: 0},
           'FTYKHX': {8: 2,9: 2,  10: 2,  11: 2,  12: 0,  23: 2,  24: 2,  28: 0,  29: 0},
           'GRKHXX': {10: 2,11: 2,  14: 2,  15: 2,  19: 0,  20: 0},
           'CLTYCK': {11: 2,12: 2,  13: 5},
           'TYCKFS': {11: 2, 12: 2,  15: 5},
           'FTYDWC': {13: 2,14: 2,  15: 5},
           'DWCKFS': {14: 2,15: 2,  16: 5},
           'CLGRCK': {12: 2,13: 2,  14: 5},
           'GRCKFS': {13: 2,14: 2,  15: 5},
           'CLZQTZ': {8: 2,9: 2,  13: 5},
           'ZQTZFS': {11: 5, 20: 2,  21: 2},
           'CLZQFX': {6: 0,8: 2,  9: 2,  10: 2,  11: 2,  16: 5},
           'ZQFXFS': {6: 0,8: 2,  9: 2,  10: 2,  11: 2,  16: 5},
           'CLGQTZ': {9: 2,10: 2},
           'GQTZFS': {10: 2, 11: 2},
           'SPVTZX': {12: 2,13: 2},
           'SPVFSX': {13: 2,14: 2},
           'WQYBJY': {7: 2}
           }


# dcm_lst = {'CLDWDK': {19: 2,20: 2,  22: 5,  24: 5},
#            'DWDKFS': {19: 2,20: 2,  22: 5,  24: 5},
#            'DBHTXX': {10: 2,11: 2,  12: 2},
#            'DBWXX': {11: 2,13: 2,  14: 2},
#            'JRJGFZ': {},  
#            'FTYKHX': {8: 2, 9: 2,  10: 2,  11: 2,  12: 0,  23: 2,  24: 2,  28: 0,  29: 0},
#            'WQYBJY': {7: 2},
#            'CLGRDK':{14:2,15:2,17:5,19:5},
#            'GRDKFS':{14:2,15:2,17:5,19:5,},
#            'GRKHXX':{10:2,11:2,14:2,15:2,19:0,20:0},
#            }


def workDirClean():
    fileList = os.listdir(workDir)
    for eachfile in fileList:
        try:
            suff = eachfile.split('.')[1]
            if suff in ['dat','log']:
                os.remove(eachfile)
        except IndexError:
            pass

def getfileNames():
    '''
    获取工作文件夹下所有的xlsx格式文件
    
    '''
    fileList = os.listdir(workDir)
    excelFiles =[]
    for eachfile in fileList:
        try:
            suff = eachfile.split('.')[1]
            if suff == 'xlsx':
                excelFiles.append(eachfile)
        except IndexError:
            pass
    return excelFiles

def xlsx2dat(fileName):
    sep = '|'
    workbook = xlrd.open_workbook(fileName)
    sheet1 = workbook.sheet_by_index(0)
    nrows = sheet1.nrows
    ncols = sheet1.ncols
    txt_file = fileName.split('.')[0] + '.dat'
    bw_type = fileName.split('_')[1]
    if dcm_lst.get(bw_type, None) == None:
        print('本程序仅支持部分数据报文的生产操作,文件列表如下：')
        print(list(dcm_lst.keys()))
        print(fileName, '暂不支持报文生产操作！')
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
                    # ctype : 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
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

def getfilemd5(fileName):
    fileName = fileName.replace('xlsx','dat')
    time1 = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    org_file = open(os.getcwd() + os.sep + fileName, 'rb')
    count = len(org_file.readlines())
    myhash = hashlib.md5()
    b = open(os.getcwd() + os.sep + fileName, 'rb').read()
    myhash.update(b)
    org_file.close()
    size = os.path.getsize(fileName)
    file = fileName.split('\\')[(-1)]
    md5 = myhash.hexdigest()
    optname = fileName.split('.')[0] + '.log'
    with open((os.getcwd() + os.sep + optname), 'w', encoding='utf8') as (file):
        file.write(fileName)
        file.write('\n')
        file.write(md5)
        file.write('\n')
        file.write(str(size))
        file.write('\n')
        file.write(time1)
        file.write('\n')
        file.write(str(count))

def zipFiles(fileName):
    fileName_pre = fileName.split('.')[0]
    zipFile = zipfile.ZipFile(os.path.join(workDir,fileName_pre) +'.zip','w', zipfile.ZIP_DEFLATED)
    f_txt = fileName_pre + '.dat'
    f_log = fileName_pre + '.log'
    try:
        zipFile.write(os.path.join(workDir, f_txt), arcname=f_txt)
        zipFile.write(os.path.join(workDir, f_log), arcname=f_log)
        zipFile.close()
        print(f'{fileName}报表压缩包搞定啦！')
    except:
        print(f'{fileName}报表dat文件或log文件丢失，请尝试重新运行本程序')

def workFlow():
    fileNames = getfileNames()
    if fileNames:
        for index,fileName in enumerate(fileNames):
            print(f'-----------正在处理第【{index+1}】个报表-----------')
            print(f'根据{fileName}生成dat和log文件')
            xlsx2dat(fileName)
            getfilemd5(fileName)
            zipFiles(fileName)
    else:
        print('当前文件夹下未找到xlsx格式报表，请检查。')


if __name__ == '__main__':
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
    os.system('pause')
    workDirClean()
    workFlow()
    workDirClean()
    os.system('pause')
