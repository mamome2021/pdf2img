#!/usr/bin/env python3
#第7版
import sys
import os
import traceback
import fitz
from PIL import Image
from io import BytesIO
import math

def read_config():
    config = {'error': False,
              'single-image': False,
              'no-crop': False,
              'prefer-png': False,
              'tiff-compression': 'packbits'}
    try:
        config_filename = os.path.abspath(os.path.join(os.path.dirname(__file__), 'config-pdf2img.txt'))
        config_file = open(config_filename, 'r', encoding='utf-8')
        lines = config_file.read().split('\n')
        for line in lines:
            option = line.split()
            if len(option) == 0:
                continue
            if option[0] == 'error':
                config['error'] = True
            elif option[0] == 'single-image':
                config['single-image'] = True
            elif option[0] == 'no-crop':
                config['no-crop'] = True
            elif option[0] == 'prefer-png':
                config['prefer-png'] = True
            elif option[0] == 'tiff-compression':
                config['tiff-compression'] = option[1]
    except FileNotFoundError:
        pass
    except Exception:
        print(traceback.format_exc())
    return config

def find_largest_image(images):
    size = 0
    index = 0
    for i, image in enumerate(images):
        if image[2] * image[3] > size:
            size = image[2] * image[3]
            index = i
    return images[index]

def fix_transparent_area(img):
    g, a = img.split()
    # 半透明的地方會變灰，因此改為全透明
    a = a.point(lambda i: i > 254 and 255)
    return Image.merge('LA', (g, a))

def render_image(page, zoom, colorspace='GRAY', alpha=True):
    pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=colorspace, alpha=alpha)
    if not alpha:
        return Image.frombytes('L', [pixmap.width, pixmap.height], pixmap.samples)
    image = Image.frombytes('LA', [pixmap.width, pixmap.height], pixmap.samples)
    return fix_transparent_area(image)

def generate_image(config, doc, page, page_noimg, image, output_dir):
    try:
        img_xref = image[0]
        width = int(doc.xref_get_key(img_xref, "Width")[1])
        height = int(doc.xref_get_key(img_xref, "Height")[1])
        cs_type = doc.xref_get_key(img_xref, "ColorSpace")[0]
        cs = doc.xref_get_key(img_xref, "ColorSpace")[1]
        output_name=f"{output_dir}/{page.number+1}-{img_xref}"
        if doc.xref_get_key(img_xref, "Filter")[1] == '/DCTDecode':
            print(output_name, "jpeg")
            pil_image = Image.open(BytesIO(doc.xref_stream_raw(img_xref)))
        elif doc.xref_get_key(img_xref,"ImageMask")[1] == 'true' or doc.xref_get_key(img_xref, "BitsPerComponent")[1] == '1':
            print(output_name, "mono")
            pil_image = Image.frombytes('1', (width,height), doc.xref_stream(img_xref))
            pil_image = pil_image.convert('L')
        elif cs_type == 'xref':
            print(output_name, "xref cs")
            # 太難了不會做，用第一版的方法
            img_dict = doc.extract_image(img_xref)
            img_data = img_dict["image"]
            pil_image = Image.open(BytesIO(img_data))
        elif cs == "/DeviceCMYK":
            print(output_name, "cmyk")
            pil_image = Image.frombytes('CMYK', (width, height), doc.xref_stream(img_xref))
        elif cs == "/DeviceGray":
            print(output_name, "gray")
            pil_image = Image.frombytes('L', (width, height), doc.xref_stream(img_xref))
        elif cs == "/DeviceRGB":
            print(output_name,"rgb")
            pil_image = Image.frombytes('RGB', (width, height), doc.xref_stream(img_xref))
        else:
            print(output_name,"wip:", cs)
            # 其他，還沒做，用第一版的方法
            img_dict = doc.extract_image(img_xref)
            img_data = img_dict["image"]
            pil_image = Image.open(BytesIO(img_data))
        image_matrix = page.get_image_rects(img_xref, transform=True)[0][1]
        if image_matrix[1:3] != (0, 0):
            print(output_name, '警告：圖片旋轉或歪斜，輸出將與pdf不同')
        zoom = width / image_matrix[0]
        img_noimg = render_image(page_noimg, width / image_matrix[0])
        if not config['no-crop']:
            img_merge = Image.new(pil_image.mode, (math.ceil(page.rect[2] * zoom), math.ceil(page.rect[3] * zoom)), color='white')
            img_merge.paste(pil_image, (round(image_matrix[4] * zoom), round(image_matrix[5] * zoom)))
            img_merge.paste(img_noimg, (0, 0), img_noimg)
        else:
            image_rect = page.get_image_rects(img_xref)[0]
            width_merge = max(page.rect[2], image_rect[2]) - min(page.rect[0], image_rect[0])
            height_merge = max(page.rect[3], image_rect[3]) - min(page.rect[1], image_rect[1])
            x_offset = min(image_rect[0], 0)
            y_offset = min(image_rect[1], 0)
            img_merge = Image.new(pil_image.mode, (math.ceil(width_merge * zoom), math.ceil(height_merge * zoom)), color='white')
            img_merge.paste(pil_image, (round(max(image_matrix[4], 0) * zoom), round(max(image_matrix[5], 0) * zoom)))
            img_merge.paste(img_noimg, (round(-x_offset * zoom), round(-y_offset * zoom)), img_noimg)
            
        return img_merge, output_name
    except Exception as e:
        print(traceback.format_exc())
        if config['error']:
            exit(1)

def main():
    config = read_config()

    for file in sys.argv[1:]:
        doc = fitz.open(file)
        doc_noimg = fitz.open(file)
        for page in doc_noimg:
            for image in page.get_images():
                xref = image[0]
                page.delete_image(xref)
        doc_noimg = fitz.open('pdf', doc_noimg.tobytes(garbage=1))
        output_dir = file + "-img"
        os.makedirs(output_dir, exist_ok=True)

        for pagenum, page in enumerate(doc):
            page_noimg = doc_noimg[pagenum]
            images = page.get_images()
            if not images:
                print(f'警告：第{pagenum+1}頁沒有圖片，使用600dpi渲染')
                image = render_image(page, 600 / 72, alpha=False)
                if config['prefer-png'] == True:
                    image.save(f"{output_dir}/{pagenum+1}.png")
                else:
                    image.save(f"{output_dir}/{pagenum+1}.tiff", compression=config['tiff-compression'])
                continue
            if len(images) > 1:
                print(f'警告：第{pagenum+1}頁包含多張圖片，輸出圖片只會包含一張圖片')
                if config['single-image']:
                    images = [find_largest_image(images)]
            for image in images:
                img_generated, output_name = generate_image(config, doc, page, page_noimg, image, output_dir)
                if config['prefer-png'] == True and img_generated.mode != 'CMYK':
                    img_generated.save(f"{output_name}.png")
                else:
                    img_generated.save(f"{output_name}.tiff", compression=config['tiff-compression'])

main()
