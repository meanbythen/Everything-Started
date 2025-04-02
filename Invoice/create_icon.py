from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    # 创建一个 256x256 的图像，使用RGBA模式（支持透明度）
    size = (256, 256)
    image = Image.new('RGBA', size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # 创建一个圆形背景
    circle_color = (52, 152, 219)  # 使用一个好看的蓝色
    circle_bbox = [20, 20, 236, 236]  # 留出一些边距
    draw.ellipse(circle_bbox, fill=circle_color)

    # 添加文字
    try:
        # 尝试使用微软雅黑字体
        font = ImageFont.truetype("msyh.ttc", 100)
    except:
        # 如果找不到，使用默认字体
        font = ImageFont.load_default()

    text = "发票"
    # 获取文字大小
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]

    # 计算文字位置，使其居中
    x = (size[0] - text_width) // 2
    y = (size[1] - text_height) // 2

    # 绘制文字
    draw.text((x, y), text, fill="white", font=font)

    # 保存为多种尺寸的ICO文件
    image.save("icon.ico", format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])

if __name__ == "__main__":
    create_icon() 