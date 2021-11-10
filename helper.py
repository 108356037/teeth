#!/usr/bin/python3
# -- coding:UTF-8 --
import xlsxwriter
import re
import os
from datetime import datetime
import shutil

image_dir = str(input("輸入照片所在的資料夾名稱："))
image_dir = '/Users/henry/Desktop/' + image_dir
now = datetime.now()
dt_string = now.strftime("%Y-%m-%d %H:%M:%S")
xlxs_name = dt_string+'-牙齒筆記.xlsx'
workbook = xlsxwriter.Workbook(xlxs_name)

pattern = re.compile("(截圖|20).*")
unsort_files_with_ds = [x for x in os.listdir(image_dir)]
unsort_files = list(filter(pattern.match, unsort_files_with_ds))
print(unsort_files)
for i, f in enumerate(unsort_files):
    unsort_files[i] = f.lstrip('截圖 ').replace(
        '下午', 'PM-').replace('上午', 'AM-').rstrip('.png').rstrip('拷貝')
    os.rename(image_dir+'/'+f, image_dir+'/' + unsort_files[i]+'.png')



sort_datetime = sorted([datetime.strptime(
    ts, "%Y-%m-%d %p-%I.%M.%S") for ts in unsort_files])
# print(sort_datetime)
sort_png_time = [date.strftime('%Y-%m-%d %p-%-I.%M.%S')
                 for date in sort_datetime]

# files_path = [image_dir+'/'+x for x in sort_png_time]
# print(files_path)

for i, img in enumerate(sort_png_time):
    worksheet = workbook.add_worksheet(img)
    worksheet.write('A1', img)
    worksheet.insert_image('A2', image_dir+'/'+img+'.png')

workbook.close()

shutil.move('/Users/henry/'+xlxs_name, '/Users/henry/Desktop/'+xlxs_name)
