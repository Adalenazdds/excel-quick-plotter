from PIL import Image

# 打开你生成的 PNG 图片
img = Image.open('icon.png')
# 保存为 ICO 格式，包含多种尺寸以适应 Windows 缩放
img.save('icon.ico', format='ICO', sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("图标转换完成！")