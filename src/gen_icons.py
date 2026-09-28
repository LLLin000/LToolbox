import os
from PIL import Image, ImageDraw

# 图标输出目录：与 engine.py / customUI14.xml 同级的 src/icons
base = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icons')
os.makedirs(base, exist_ok=True)

def icon(name, draw_fn):
    im = Image.new('RGBA', (32, 32), (0,0,0,0))
    d = ImageDraw.Draw(im)
    draw_fn(d)
    p = os.path.join(base, name + '.png')
    im.save(p)
    return p

# 1. 帮助：蓝色圆底 + 白色问号
def d_help(d):
    d.ellipse([2,2,29,29], fill='#2B579A')
    d.arc([11,7,21,17], start=210, end=30, fill='white', width=3)
    d.line([16,17,16,21], fill='white', width=3)
    d.point([16,25], fill='white'); d.ellipse([14.5,23.5,17.5,26.5], fill='white')
icon('help', d_help)

# 2. 标定比例尺：直尺（带刻度）
def d_calib(d):
    d.rectangle([4,10,27,22], fill='#F0F0F0', outline='#2B579A', width=2)
    for x in [8, 12, 16, 20, 24]:
        h = 5 if x in (8, 16, 24) else 3
        d.line([x, 10, x, 10+h], fill='#2B579A', width=1)
icon('calib', d_calib)

# 3. 生成取景框：红色方框 + 4 个角点
def d_makebox(d):
    d.rectangle([5,5,26,26], outline='#D83B01', width=2)
    for x,y in [(5,5),(26,5),(5,26),(26,26)]:
        d.rectangle([x-2,y-2,x+2,y+2], fill='#D83B01')
icon('makebox', d_makebox)

# 4. 加比例尺：一条黑色粗标尺 + 上方两根垂直线
def d_scalebar(d):
    # 底图缩略
    d.rectangle([3,5,28,27], outline='#CCCCCC', width=1)
    # 比例尺黑条
    d.rectangle([7,21,24,24], fill='#000000')
    # 两端竖线
    d.line([7,18,7,24], fill='#000000', width=2)
    d.line([24,18,24,24], fill='#000000', width=2)
    # 顶部示意数字
    d.line([13,14,18,14], fill='#2B579A', width=2)
icon('scalebar', d_scalebar)

# 5. 批量原图裁取：源图内多框 → 多张小图
def d_batchcrop(d):
    d.rectangle([2, 5, 20, 27], outline='#2B579A', width=2)
    d.rectangle([5, 8, 11, 14], outline='#D83B01', width=2)
    d.rectangle([12, 17, 18, 24], outline='#D83B01', width=2)
    d.polygon([(19, 13), (24, 16), (19, 19)], fill='#107C41')
    d.rectangle([24, 6, 30, 13], outline='#D83B01', width=2)
    d.rectangle([24, 20, 30, 27], outline='#D83B01', width=2)
icon('batchcrop', d_batchcrop)

print('icons written to', base)
