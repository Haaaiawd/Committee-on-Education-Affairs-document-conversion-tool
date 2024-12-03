from PIL import Image, ImageDraw

# 创建一个128x128的图像，背景白色
img = Image.new('RGB', (128, 128), 'white')
draw = ImageDraw.Draw(img)

# 画一个简单的文档图标
draw.rectangle([30, 20, 98, 108], outline='blue', width=2)
draw.polygon([(98, 20), (98, 40), (78, 20)], fill='blue')

# 保存为ICO文件
img.save('app.ico', format='ICO') 