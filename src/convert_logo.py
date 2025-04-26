from PIL import Image
import os

def convert_png_to_ico():
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    
    # 输入和输出路径
    input_path = os.path.join(project_root, 'resources', 'logo.png')
    output_path = os.path.join(project_root, 'resources', 'logo.ico')
    
    # 确保resources目录存在
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 转换图片
    img = Image.open(input_path)
    
    # 调整大小，确保是正方形
    size = max(img.size)
    new_img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    new_img.paste(img, ((size - img.size[0]) // 2, (size - img.size[1]) // 2))
    
    # 保存为ICO
    new_img.save(output_path, format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])

if __name__ == '__main__':
    convert_png_to_ico() 