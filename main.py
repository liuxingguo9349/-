import os
from aip import AipOcr
import re
from pandas import DataFrame
from easygui import diropenbox
from tqdm import tqdm

print('程序开始运行，进度如下：')
APP_ID = '25937986'
API_KEY = 'LbwSmzBAt5HpLf4IPfmhRGoE'
SECRET_KEY = 'rjuOCgwiBRoDX0qgyXhmurPa3RVA2FTE'
client = AipOcr(APP_ID, API_KEY, SECRET_KEY)


def get_file_content(filepath):
    with open(filepath, 'rb') as fp:
        return fp.read()


path = diropenbox('请选择需要处理的文件夹')
filelist = os.listdir(path)
result = []
fileList = []
fileName = ''
for file in tqdm(filelist):
    if fileName[0:14] == file[0:14]:
        fileList.append(fileName)
    fileName = file
for file in tqdm(fileList):
    filelist.remove(file)
for file in tqdm(filelist):
    a = {}
    file_path = f'{path}\\{os.path.basename(file)}'
    img = get_file_content('{}\\'.format(path) + file)
    message = client.basicGeneral(img)['words_result']
    mes = ';'.join([str(list(i.values())[0]) for i in message])
    time = re.compile(r'..-.....\d....').findall(mes)
    t = ''.join(time)
    date = t[0:5]
    times = t[5:13]
    color = ''.join(re.compile(r';.码').findall(mes))[1:3]
    a['颜色'] = color
    a['截图日期'] = date
    a['截图时间'] = times
    a['文件名'] = file
    result.append(a)
DataFrame(result).to_excel('result.xlsx')
os.startfile('result.xlsx')
input('程序运行完成')
