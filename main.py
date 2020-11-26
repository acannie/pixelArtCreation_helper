from PIL import Image
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter
import math
import pandas as pd
from sklearn.cluster import KMeans
import colorsys

# RGBの数値が1桁の場合冒頭に0をつけて2桁にして返す


def addZero(hexRGB):
    if len(hexRGB) == 3:
        return '0x0' + hexRGB[2]
    else:
        return hexRGB

# RGBをカラーコードに変換


def RGBtoColorCode(R, G, B):
    color_code = '{}{}{}'.format(addZero(hex(R)), addZero(
        hex(G)), addZero(hex(B)))  # 0xAA0xBB0xCC の形
    color_code = color_code.replace('0x', '')
    return color_code

# 色相 (0 < H < 360)
def getHue(RGBlist=[0,0,0]):
    max_color = max(RGBlist)
    min_color = min(RGBlist)
    r = RGBlist[0]
    g = RGBlist[1]
    b = RGBlist[2]
    H = 0
    if r > g and r > b:
        H = 60 * ((g - b) / (max_color - min_color))
    elif g > b and g > r:
        H = 60 * ((b - r) / (max_color - min_color)) + 120
    elif b > r and b > g:
        H = 60 * ((r - g) / (max_color - min_color)) + 240
    else:
        H = 0

    if H < 0:
        H += 360
    
    return H

# 彩度 (0 < S < 1)
def getSaturation(RGBlist=[0,0,0]):
    max_color = max(RGBlist)
    min_color = min(RGBlist)
    S = (max_color - min_color) / max_color
    return S

# 明度 (0 < V < 255)
def getBrightness(RGBlist=[0,0,0]):
    max_color = max(RGBlist)
    V = max_color
    return V

# 背景色によって文字色を変える
def getFontColor(RGBlist=[0,0,0]):
    if getBrightness(RGBlist) > 200:
        return '000000'
    else:
        return 'FFFFFF'



# あつ森の仕様
art_wh = 32
art_palette = 15
art_color_variety = 30
art_brightness_variety = 15
art_vividness_variety = 15

# ファイルのインポート
input_file = 'icon_akane.jpg'
im = np.array(Image.open(input_file))

# 画像のサイズを取得
originalImg_height = im.shape[0]
originalImg_width = im.shape[1]

# 1辺の長さがart_whの倍数になるよう正方形にトリミング（中央寄せ）
short_side_len = min(originalImg_height, originalImg_width)
long_side_len = max(originalImg_height, originalImg_width)
shard = short_side_len % art_wh
target_size = short_side_len - shard
im = np.delete(im, np.s_[short_side_len - math.floor(shard / 2):], 0)
im = np.delete(im, np.s_[:math.ceil(shard / 2)], 0)
im = np.delete(im, np.s_[long_side_len -
                         math.floor((long_side_len - target_size) / 2):], 1)
im = np.delete(im, np.s_[:math.ceil((long_side_len - target_size)/2)], 1)

img_new = Image.fromarray(im.astype(np.uint8))
img_new.save('square_' + input_file)

# 画像のサイズを取得
img_height = im.shape[0]
img_width = im.shape[1]


# ワークブックを作成
wb = openpyxl.Workbook()
ws = wb.worksheets[0]


# デザインゾーンの枠線を追加
side1 = Side(style='thick', color='000000')
for i in range(3, art_wh + 1):
    ws.cell(row=2, column=i).border = Border(top=side1)
for i in range(3, art_wh + 1):
    ws.cell(row=i, column=2).border = Border(left=side1)
for i in range(3, art_wh + 1):
    ws.cell(row=i, column=art_wh+1).border = Border(right=side1)
for i in range(3, art_wh + 1):
    ws.cell(row=art_wh+1, column=i).border = Border(bottom=side1)
ws.cell(row=2, column=2).border = Border(top=side1, left=side1)
ws.cell(row=2, column=art_wh+1).border = Border(top=side1, right=side1)
ws.cell(row=art_wh+1, column=2).border = Border(left=side1, bottom=side1)
ws.cell(row=art_wh+1, column=art_wh +
        1).border = Border(right=side1, bottom=side1)

# デザインゾーンのセルを正方形に整形
for x in range(1, art_wh+1):
    ws.column_dimensions[get_column_letter(x+1)].width = 4
for y in range(1, art_wh+1):
    ws.row_dimensions[y+1].height = 22


# art_wh * art_wh のドット絵を生成
RGB_art_wh = []
df_all_color = pd.DataFrame(columns=['R', 'G', 'B'])

color_count = 0

for r in range(art_wh):
    RGB_art_wh.append([])

    for c in range(art_wh):
        RGB_ave = [0, 0, 0]

        # n×nピクセルの平均RGBをセルの背景色とする
        n = int(img_width / art_wh)
        for i in range(n):
            for j in range(n):
                for color in range(im.shape[2]):
                    RGB_ave[color] += im[n * r + i][n * c + j][color]

        for color in range(im.shape[2]):
            RGB_ave[color] = round(RGB_ave[color] / n**2)

        RGB_art_wh[r].append(RGB_ave)
        df_all_color = df_all_color.append(
            {'R': RGB_ave[0], 'G': RGB_ave[1], 'B': RGB_ave[2]}, ignore_index=True)
        color_count += 1

cust_array = np.array([df_all_color['R'].tolist(),
                       df_all_color['G'].tolist(),
                       df_all_color['B'].tolist(), ], np.int32)

# 行列を転置
cust_array = cust_array.T

# クラスタ分析(クラスタ数=15)で色の種類を絞る
pred = KMeans(n_clusters=15).fit_predict(cust_array)
df_all_color['color'] = pred

cluster_RGB_list = []
HSV_list = []
for i in range(15):
    cluster_RGB = df_all_color[df_all_color['color'] == i].mean()

    cluster_RGB_list.append([])
    cluster_RGB_list[i].append(round(cluster_RGB[0]))
    cluster_RGB_list[i].append(round(cluster_RGB[1]))
    cluster_RGB_list[i].append(round(cluster_RGB[2]))
    
    # HSV_list[i].append(round(getHue(list(cluster_RGB))/360 * 30))
    # HSV_list[i].append(round(getSaturation(list(cluster_RGB))*15))
    # HSV_list[i].append(round(getBrightness(list(cluster_RGB))/255*15))

    hsv = colorsys.rgb_to_hsv(cluster_RGB[0], cluster_RGB[1], cluster_RGB[2])
    HSV_list.append([])
    HSV_list[i].append(math.ceil(hsv[0] * 30))
    HSV_list[i].append(math.ceil(hsv[1] * 15))
    HSV_list[i].append(math.ceil(hsv[2]/255 * 15))
    # HSV_list[i].append(round(getSaturation(list(cluster_RGB))*15))
    # HSV_list[i].append(round(getBrightness(list(cluster_RGB))/255*15))

print(HSV_list)

# 描画
for index, row in df_all_color.iterrows():
    r = cluster_RGB_list[row['color']][0]
    g = cluster_RGB_list[row['color']][1]
    b = cluster_RGB_list[row['color']][2]
    color_code = RGBtoColorCode(r, g, b)
    fill = PatternFill(patternType='solid', fgColor=color_code)
    row_num = math.ceil((index + 1) / art_wh) + 1
    col_num = index % art_wh + 2
    ws.cell(row=row_num, column=col_num).fill = fill
    ws.cell(row=row_num, column=col_num).value = row['color'] + 1
    ws.cell(row=row_num, column=col_num).font = openpyxl.styles.fonts.Font(color=getFontColor([r,g,b]))



# カラーパレット
ws['AJ2'].value = '色相'
ws['AK2'].value = '彩度'
ws['AL2'].value = '明度'

for i in range(art_palette):
    r = cluster_RGB_list[i][0]
    g = cluster_RGB_list[i][1]
    b = cluster_RGB_list[i][2]
    color_code = RGBtoColorCode(r, g, b)
    fill = PatternFill(patternType='solid', fgColor=color_code)
    ws.cell(row=i+3, column=art_wh + 3).fill = fill
    ws.cell(row=i+3, column=art_wh + 3).value = i + 1
    ws.cell(row=i+3, column=art_wh + 3).font = openpyxl.styles.fonts.Font(color=getFontColor([r,g,b]))

    ws.cell(row=i+3, column=art_wh + 4).value = HSV_list[i][0]
    ws.cell(row=i+3, column=art_wh + 5).value = HSV_list[i][1]
    ws.cell(row=i+3, column=art_wh + 6).value = HSV_list[i][2]

# ファイルのエクスポート
print('')
print("saving...")

output_file = input_file.replace('.jpg', '.xlsx')
wb.save(output_file)
wb.close()

print("complete!")
