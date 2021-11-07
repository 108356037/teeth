# -- coding:UTF-8 --
import xlsxwriter

import os
from datetime import datetime

image_dir = str(input("輸入照片所在的資料夾名稱："))
image_dir = './' + image_dir
now = datetime.now()
dt_string = now.strftime("%Y-%m-%d %H:%M:%S")
xlxs_name = dt_string+'-牙齒筆記.xlsx'
workbook = xlsxwriter.Workbook(xlxs_name)

unsort_files = [x for x in os.listdir(image_dir)]
for i, f in enumerate(unsort_files):
    unsort_files[i] = f.lstrip('截圖 ').replace(
        '下午', 'PM-').replace('上午', 'AM-').rstrip('.png')
    os.rename(image_dir+'/'+f, image_dir+'/' + unsort_files[i]+'.png')

# print(unsort_files)
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
