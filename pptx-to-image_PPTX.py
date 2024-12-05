import os          # 用于文件和路径操作
import sys         # 用于系统功能和退出程序
from win32com import client  # 用于与 Windows COM 接口交互，控制 PowerPoint
from win32com.client import constants
from PIL import Image        # 用于图像处理
import re          # 用于正则表达式处理文件名

def pptx_to_images(pptx_path, image_folder, dpi=300):
    """
    将 PPTX 文件的每一张幻灯片导出为高分辨率的 PNG 图像。

    参数：
        pptx_path (str): PPTX 文件的路径。
          (str): 导出图像的保存文件夹路径。
        dpi (int): 导出图像的分辨率（每英寸点数），默认值为 300 DPI。
    """
    # 创建 PowerPoint 应用程序的 COM 对象
    powerpoint = client.Dispatch('PowerPoint.Application')
    powerpoint.Visible = 1  # 将 PowerPoint 应用程序设置为可见（1）或不可见（0）

    # 打开指定的 PPTX 文件
    ppt = powerpoint.Presentations.Open(pptx_path)
    
    # 如果图像保存文件夹不存在，则创建它
    if not os.path.exists(image_folder):
        os.makedirs(image_folder)
    
    # 获取幻灯片的宽度和高度，单位为点（Points）
    slide_width = ppt.PageSetup.SlideWidth
    slide_height = ppt.PageSetup.SlideHeight
    
    # 将幻灯片尺寸从点转换为像素尺寸
    # 1英寸 = 72点，因此英寸 = 点 / 72
    # 然后像素尺寸 = 英寸 * DPI
    scale_width = int((slide_width / 72) * dpi)
    scale_height = int((slide_height / 72) * dpi)

    # 遍历每一张幻灯片，导出为指定尺寸的 PNG 图像
    for i, slide in enumerate(ppt.Slides):
        # 构造图像文件的保存路径，命名为 Slide_1.png, Slide_2.png, 等
        image_path = os.path.join(image_folder, f"Slide_{i+1}.png")
        # 导出幻灯片为图像文件
        slide.Export(image_path, "PNG", scale_width, scale_height)
    
    # 关闭演示文稿和 PowerPoint 应用程序
    ppt.Close()
    powerpoint.Quit()
    
    # 返回幻灯片的尺寸
    return slide_width, slide_height

def images_to_pptx(image_folder, pptx_path, slide_width, slide_height):
    """
    将指定文件夹中的所有 PNG 图像插入到一个新的 PPTX 文件中，每张图像作为一张幻灯片。
    
    参数：
        image_folder (str): 存放 PNG 图像的文件夹路径。
        pptx_path (str): 输出的 PPTX 文件路径。
        slide_width (float): 幻灯片宽度（点）
        slide_height (float): 幻灯片高度（点）
    """
    # 获取图像文件夹中所有的 PNG 文件列表
    image_files = [os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.endswith('.png')]
    if not image_files:
        print("没有找到任何 PNG 图像文件。")
        sys.exit(1)
    
    # 定义一个函数，从文件名中提取幻灯片编号
    def get_slide_number(filename):
        # 使用正则表达式匹配文件名中的数字部分
        match = re.search(r'Slide_(\d+)\.png', filename)
        if match:
            return int(match.group(1))  # 返回数字部分，作为幻灯片编号
        else:
            return 0  # 如果未匹配到，返回 0，将其排在前面
    
    # 按照幻灯片编号对图像文件进行排序
    image_files.sort(key=get_slide_number)
    
    # 创建 PowerPoint 应用程序的 COM 对象
    powerpoint = client.Dispatch('PowerPoint.Application')
    powerpoint.Visible = 1  # 可见模式
    
    # 创建一个新的演示文稿
    ppt = powerpoint.Presentations.Add()
    
    # 设置幻灯片尺寸
    ppt.PageSetup.SlideWidth = slide_width
    ppt.PageSetup.SlideHeight = slide_height

    ppLayoutBlank = 12  # 定义 ppLayoutBlank 常量的值
    # 遍历所有的图像文件，添加到幻灯片中
    for image_file in image_files:
        # 添加一张空白幻灯片
        slide = ppt.Slides.Add(ppt.Slides.Count + 1, ppLayoutBlank)   
        # 在幻灯片中添加图片
        slide.Shapes.AddPicture(FileName=image_file, LinkToFile=False, SaveWithDocument=True, Left=0, Top=0, Width=slide_width, Height=slide_height)
    
    # 保存新的 PPTX 文件
    ppt.SaveAs(pptx_path)
    # 关闭演示文稿和 PowerPoint 应用程序
    ppt.Close()
    powerpoint.Quit()

def main():
    """
    主函数，负责与用户交互，获取输入，并调用其他函数完成转换过程。
    """
    # 提示用户输入要转换的 PPTX 文件路径，并去除两端的引号
    pptx_path = input("请输入要转换的 .pptx 文件路径：").strip('"')
    pptx_path = os.path.abspath(pptx_path)  # 将路径转换为绝对路径
    # 检查文件是否存在
    if not os.path.exists(pptx_path):
        print("文件不存在。")
        sys.exit(1)
    
    # 分割文件路径，获取目录和文件名
    pptx_dir, pptx_file = os.path.split(pptx_path)
    # 分割文件名和扩展名，获取文件名前缀
    pptx_name, _ = os.path.splitext(pptx_file)
    # 设置图像保存文件夹路径，命名为 <文件名>_images
    image_folder = os.path.join(pptx_dir, pptx_name + '_images')
    
    # 动态调整分辨率
    dpi_input = input("请输入所需的分辨率（DPI），直接回车则使用默认值 300 DPI：")
    if dpi_input.strip() == '':
        desired_dpi = 300  # 默认分辨率为 300 DPI
    else:
        try:
            desired_dpi = int(dpi_input)
            if desired_dpi <= 0:
                print("分辨率必须是正整数。使用默认值 300 DPI。")
                desired_dpi = 300
        except ValueError:
            print("输入无效。使用默认值 300 DPI。")
            desired_dpi = 300
    
    # 调用函数，将 PPTX 转换为图像文件，并获取幻灯片尺寸
    slide_width, slide_height = pptx_to_images(pptx_path, image_folder, dpi=desired_dpi)
    
    # 设置输出的 PPTX 文件路径，与原文件同名，加后缀 _图片版.pptx
    output_pptx_path = os.path.join(pptx_dir, pptx_name + '_图片版.pptx')
    # 调用函数，将图像文件插入到新的 PPTX 文件
    images_to_pptx(image_folder, output_pptx_path, slide_width, slide_height)
    
    # 删除中间的图像文件
    import shutil
    shutil.rmtree(image_folder)
    
    print(f"图片版 PPTX 文件已保存到：{output_pptx_path}")

if __name__ == '__main__':
    main()
