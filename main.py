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
import sys

def trimmingImg(im=[[[0,0,0]]]):
    IMG_HEIGHT = im.shape[0]
    IMG_WIDTH = im.shape[1]

    LONG_SIDE_AXIS = None
    SHORT_SIDE_AXIS = None

    if IMG_HEIGHT < IMG_WIDTH: # 画像が横長
        LONG_SIDE_AXIS = 1
        SHORT_SIDE_AXIS = 0
    else: # 画像が縦長
        LONG_SIDE_AXIS = 0
        SHORT_SIDE_AXIS = 1
    
    LONG_SIDE_LENGTH = max(IMG_HEIGHT, IMG_WIDTH)
    SHORT_SIDE_LENGTH = min(IMG_HEIGHT, IMG_WIDTH)

    TARGET_LENGTH = SHORT_SIDE_LENGTH - SHORT_SIDE_LENGTH % ART_WH # 1辺の長さが ART_WIDTH の倍数になるよう切り落とし
    SHARD = LONG_SIDE_LENGTH - TARGET_LENGTH

    im = np.delete(im, np.s_[SHORT_SIDE_LENGTH - math.floor((SHORT_SIDE_LENGTH - TARGET_LENGTH) / 2):], SHORT_SIDE_AXIS)
    im = np.delete(im, np.s_[:math.ceil((SHORT_SIDE_LENGTH - TARGET_LENGTH) / 2)], SHORT_SIDE_AXIS)
    im = np.delete(im, np.s_[LONG_SIDE_LENGTH - math.floor(SHARD / 2):], LONG_SIDE_AXIS)
    im = np.delete(im, np.s_[:math.ceil(SHARD / 2)], LONG_SIDE_AXIS)

    return im

# RGBの数値が1桁の場合冒頭に0をつけて2桁にして返す
def addZero(strHexNum='0xAA'):
    if len(strHexNum) == 3:
        return '0x0' + strHexNum[2]
    elif len(strHexNum) == 4:
        return strHexNum
    else:
        print('error!')
        sys.exit()

# RGBをカラーコードに変換
def RGBtoColorCode(RGB=[0, 0, 0]):
    color_code = ''
    for color in range(COLOR_VARIETY):
        color_code += str(addZero(hex(RGB[color])))
    color_code = color_code.replace('0x', '') # '0xrr0xgg0xbb' を 'rrggbb' にする
    return color_code


# 明度 (0 < V < 255)
def getBrightness(RGB=[0,0,0]):
    max_color = max(RGB)
    V = max_color
    return V

# 背景色によって文字色を変える
def getFontColor(RGB=[0,0,0]):
    if getBrightness(RGB) > 200:
        return '000000'
    else:
        return 'FFFFFF'

# デザインゾーンの枠線を追加
def drawRuledLine(ws=openpyxl.Workbook().worksheets[0]):
    thick = Side(style='thick', color='000000')
    for i in range(3, ART_WH + 1):
        ws.cell(row=2, column=i).border = Border(top=thick)
    for i in range(3, ART_WH + 1):
        ws.cell(row=i, column=2).border = Border(left=thick)
    for i in range(3, ART_WH + 1):
        ws.cell(row=i, column=ART_WH+1).border = Border(right=thick)
    for i in range(3, ART_WH + 1):
        ws.cell(row=ART_WH+1, column=i).border = Border(bottom=thick)
    ws.cell(row=2, column=2).border = Border(top=thick, left=thick)
    ws.cell(row=2, column=ART_WH+1).border = Border(top=thick, right=thick)
    ws.cell(row=ART_WH+1, column=2).border = Border(left=thick, bottom=thick)
    ws.cell(row=ART_WH+1, column=ART_WH +
            1).border = Border(right=thick, bottom=thick)

    # デザインゾーンの破線を追加
    mediumDashed = Side(style='mediumDashed', color='000000')
    for i in range(3, ART_WH + 1):
        ws.cell(row=round(ART_WH / 2) + 1, column=i).border = Border(bottom=mediumDashed)
    for i in range(3, ART_WH + 1):
        ws.cell(row=i, column=round(ART_WH / 2) + 1).border = Border(right=mediumDashed)
    ws.cell(row=round(ART_WH / 2) + 1, column=round(ART_WH / 2) + 1).border = Border(right=mediumDashed, bottom=mediumDashed)
    ws.cell(row=round(ART_WH / 2) + 2, column=round(ART_WH / 2) + 2).border = Border(top=mediumDashed, left=mediumDashed)
    ws.cell(row=2, column=round(ART_WH / 2) + 1).border = Border(top=thick, right=mediumDashed)
    ws.cell(row=round(ART_WH / 2) + 1, column=2).border = Border(left=thick, bottom=mediumDashed)
    ws.cell(row=round(ART_WH / 2) + 1, column=ART_WH + 1).border = Border(right=thick, bottom=mediumDashed)
    ws.cell(row=ART_WH + 1, column=round(ART_WH / 2) + 1).border = Border(bottom=thick, right=mediumDashed)

# デザインゾーンのセルを正方形に整形
def makeCellsSquare(ws=openpyxl.Workbook().worksheets[0]):
    for x in range(1, ART_WH+1):
        ws.column_dimensions[get_column_letter(x+1)].width = 4
    for y in range(1, ART_WH+1):
        ws.row_dimensions[y+1].height = 22

# 定数
R = 0
G = 1
B = 2

# あつ森の仕様
ART_WH = 32
ART_PALETTE = 15
ART_COLOR_VARIETY = 30
ART_BRIGHTNESS_VARIERY = 15
ART_VIVIDNESS_VARIETY = 15

# ファイルのインポート
if len(sys.argv) != 2:
    print('please specify the image file.')
    sys.exit()

input_file = sys.argv[1]
im = np.array(Image.open('figure/' + input_file))

im = trimmingImg(im)

# 切り取った画像を保存
Image.fromarray(im.astype(np.uint8)).save('squaredFigure/' + input_file)

# 画像のサイズを取得
IMG_HEIGHT = im.shape[0]
IMG_WIDTH = im.shape[1]
COLOR_VARIETY = im.shape[2]

# ワークブックを作成
wb = openpyxl.Workbook()
ws = wb.worksheets[0]

drawRuledLine(ws)
makeCellsSquare(ws)

# ART_WH * ART_WH のドット絵を生成
RGB_art_wh = []
df_all_color = pd.DataFrame(columns=['R', 'G', 'B'])

color_count = 0

for r in range(ART_WH):
    RGB_art_wh.append([])

    for c in range(ART_WH):
        # N x N ピクセルの平均 RGB をセルの背景色とする
        ave_RGB = [0, 0, 0]

        N = math.floor(IMG_WIDTH / ART_WH)
        for i in range(N):
            for j in range(N):
                for color in range(COLOR_VARIETY):
                    ave_RGB[color] += im[N * r + i][N * c + j][color]

        for color in range(COLOR_VARIETY):
            ave_RGB[color] = round(ave_RGB[color] / N**2)

        RGB_art_wh[r].append(ave_RGB)
        df_all_color = df_all_color.append(
            {'R': ave_RGB[0], 'G': ave_RGB[1], 'B': ave_RGB[2]}, ignore_index=True)
        color_count += 1

color_array = np.array([df_all_color['R'].tolist(),
                       df_all_color['G'].tolist(),
                       df_all_color['B'].tolist(), ], np.int32)

# 行列を転置
color_array = color_array.T

# クラスタ分析(クラスタ数=ART_PALETTE)で色の種類を絞る
pred = KMeans(n_clusters=ART_PALETTE).fit_predict(color_array)

# 分類ナンバーを追加
df_all_color['color'] = pred

cluster_RGB_list = []
cluster_HSV_list = []
for i in range(ART_PALETTE):
    # クラスターの平均色を取得
    cluster_RGB = df_all_color[df_all_color['color'] == i].mean()
    
    cluster_RGB_list.append([])
    
    for color in cluster_RGB:
        cluster_RGB_list[i].append(round(color))
    
    hsv = colorsys.rgb_to_hsv(cluster_RGB[0], cluster_RGB[1], cluster_RGB[2])
    cluster_HSV_list.append([])

    cluster_HSV_list[i].append(math.ceil(hsv[0] * ART_COLOR_VARIETY))
    cluster_HSV_list[i].append(math.ceil(hsv[1] * ART_BRIGHTNESS_VARIERY))
    cluster_HSV_list[i].append(math.ceil(hsv[2]/255 * ART_VIVIDNESS_VARIETY))
    
    # 上記では、いずれも要素の値が0より大きいと仮定して切り上げているため、0の場合は1に変更する
    for j in range(COLOR_VARIETY):
        if cluster_HSV_list[i][j] == 0:
            cluster_HSV_list[i][j] = 1


# 描画
for index, row in df_all_color.iterrows():
    color_code = RGBtoColorCode(cluster_RGB_list[row['color']])
    fill = PatternFill(patternType='solid', fgColor=color_code)
    row_num = math.ceil((index + 1) / ART_WH) + 1
    col_num = index % ART_WH + 2
    ws.cell(row=row_num, column=col_num).fill = fill
    ws.cell(row=row_num, column=col_num).value = row['color'] + 1
    ws.cell(row=row_num, column=col_num).font = openpyxl.styles.fonts.Font(color=getFontColor(cluster_RGB_list[row['color']]))


# カラーパレット
ws['AJ2'].value = '色相'
ws['AK2'].value = '彩度'
ws['AL2'].value = '明度'

for i in range(ART_PALETTE):
    color_code = RGBtoColorCode(cluster_RGB_list[i])
    fill = PatternFill(patternType='solid', fgColor=color_code)
    ws.cell(row=i+3, column=ART_WH + 3).fill = fill
    ws.cell(row=i+3, column=ART_WH + 3).value = i + 1
    ws.cell(row=i+3, column=ART_WH + 3).font = openpyxl.styles.fonts.Font(color=getFontColor(cluster_RGB_list[i]))

    for factor in range(COLOR_VARIETY):
        ws.cell(row=i+3, column=ART_WH + 4 + factor).value = cluster_HSV_list[i][factor]

# ファイルのエクスポート
print("saving...")

output_file = input_file.replace('.jpg', '.xlsx')
wb.save('result/' + output_file)
wb.close()

print("complete!")
