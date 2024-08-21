#!/usr/bin/env python3
from io import BytesIO
from itertools import repeat
import math
from multiprocessing import Pool
import os
import sys
import traceback

import cairo
import fitz
from PIL import Image, ImageOps
import pillow_jxl


def read_config():
    config = {'processes': 2,
              'only-extract': False,
              'render-image': False,
              'no-crop': False,
              'extract-jpeg': False,
              'prefer-mono': False,
              'save-jxl': False,
              'save-png': False,
              'save-tiff': ''}
    try:
        if 'PDF2IMG_CONFIG' in os.environ:
            config_filename = os.environ['PDF2IMG_CONFIG']
        elif getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # In PyInstaller bundle
            config_filename = os.path.abspath(os.path.join(os.path.dirname(sys.executable), 'config-pdf2img.txt'))
        else:
            config_filename = os.path.abspath(os.path.join(os.path.dirname(__file__), 'config-pdf2img.txt'))
        with open(config_filename, 'r', encoding='utf-8') as config_file:
            lines = config_file.read().split('\n')
        for line in lines:
            option = line.split()
            if len(option) == 0:
                continue
            elif option[0] == 'processes':
                config['processes'] = int(option[1])
            elif option[0] == 'only-extract':
                config['only-extract'] = True
            elif option[0] == 'render-image':
                config['render-image'] = True
            elif option[0] == 'no-crop':
                config['no-crop'] = True
            elif option[0] == 'extract-jpeg':
                config['extract-jpeg'] = True
            elif option[0] == 'prefer-mono':
                config['prefer-mono'] = True
            elif option[0] == 'save-jxl':
                config['save-jxl'] = True
            elif option[0] == 'save-png':
                config['save-png'] = True
            elif option[0] == 'save-tiff':
                config['save-tiff'] = option[1]
    except FileNotFoundError:
        print('警告：找不到設定檔')
    except Exception:
        print(traceback.format_exc())
    return config

def get_referencer_of_image(image, page):
    ref = image[9]
    if ref == 0:
        # Image is directly referenced by the page
        return page.get_contents()[0]
    else:
        return ref

def remove_path_fill(doc, page):
    images = page.get_images(full=True)
    if len(images) == 0:
        return
    image = images[0]
    image_name = image[7]
    ref = get_referencer_of_image(image, page)
    stream = doc.xref_stream(ref)
    stream_split = stream.split(f'/{image_name} Do\n'.encode(), 1)
    stream_split[0] = stream_split[0].replace(b'\nf\n', b'\nn\n').replace(b'\nf*\n', b'\nn\n')
    doc.update_stream(ref, f'/{image_name} Do\n'.encode().join(stream_split))

def find_largest_image(images):
    size = 0
    index = 0
    for i, image in enumerate(images):
        if image[2] * image[3] > size:
            size = image[2] * image[3]
            index = i
    return index

def render_image(page, zoom, colorspace='GRAY', alpha=True):
    if alpha == True and colorspace == 'CMYK':
        #CMYK with alpha channel is not supported by pillow
        colorspace = 'RGB'
    if colorspace == 'L':
        colorspace = 'GRAY'
    pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=colorspace, alpha=alpha)
    if colorspace == 'GRAY':
        colorspace = 'L'
    if not alpha:
        return Image.frombytes(colorspace, (pixmap.width, pixmap.height), pixmap.samples)
    image = Image.frombytes(colorspace + 'a', (pixmap.width, pixmap.height), pixmap.samples)
    image = image.convert(colorspace + 'A')
    return image

def extract_image(doc, img_xref, pagenum_str):
    width = int(doc.xref_get_key(img_xref, "Width")[1])
    height = int(doc.xref_get_key(img_xref, "Height")[1])
    cs_type = doc.xref_get_key(img_xref, "ColorSpace")[0]
    cs = doc.xref_get_key(img_xref, "ColorSpace")[1]
    if doc.xref_get_key(img_xref, "Filter")[1] == '/DCTDecode':
        if cs == "/DeviceCMYK":
            # Using xref_stream_raw directly produces image with inverted color
            pixmap = fitz.Pixmap(doc, img_xref)
            return "cmyk", Image.frombytes('CMYK', (pixmap.width, pixmap.height), pixmap.samples)
        else:
            return "jpeg", doc.xref_stream_raw(img_xref)
    elif doc.xref_get_key(img_xref,"ImageMask")[1] == 'true':
        return "mono-mask", Image.frombytes('1', (width, height), doc.xref_stream(img_xref))
    elif doc.xref_get_key(img_xref, "BitsPerComponent")[1] == '1':
        return "mono", Image.frombytes('1', (width, height), doc.xref_stream(img_xref))
    elif cs_type == 'xref':
        print(f"警告：{pagenum_str}-{img_xref} xref cs")
        # 太難了不會做，用第一版的方法
        img_dict = doc.extract_image(img_xref)
        img_data = img_dict["image"]
        return "rgb", Image.open(BytesIO(img_data))
    elif cs == "/DeviceCMYK":
        return "cmyk", Image.frombytes('CMYK', (width, height), doc.xref_stream(img_xref))
    elif cs == "/DeviceGray":
        return "gray", Image.frombytes('L', (width, height), doc.xref_stream(img_xref))
    elif cs == "/DeviceRGB":
        return "rgb", Image.frombytes('RGB', (width, height), doc.xref_stream(img_xref))
    else:
        print(f"警告：{pagenum_str}-{img_xref}未知色彩空間", cs)
        # 其他，還沒做，用第一版的方法
        img_dict = doc.extract_image(img_xref)
        img_data = img_dict["image"]
        return "rgb", Image.open(BytesIO(img_data))

def save_extracted_image(config, doc, page, image, output_dir):
    img_xref = image[0]
    pagenum_str = str(page.number + 1).zfill(3)
    output_name = f"{output_dir}/{pagenum_str}-{img_xref}"
    image_type, image_extract = extract_image(doc, img_xref, pagenum_str)
    if image_type == 'jpeg':
        with open(f"{output_name}.jpg",'wb') as f:
            f.write(image_extract)
    else:
        save_pil_image(config, image_extract, output_name)

def create_clipping_path_image(doc, page, image, size, image_pos, image_size):
    image_name = image[7]
    image_xref = image[0]
    width = int(doc.xref_get_key(image_xref, "Width")[1])
    referencer = get_referencer_of_image(image, page)
    stream = doc.xref_stream(referencer).split(f'\n/{image_name} Do\n'.encode())[0].split(b'\nQ\n')[-1]
    if not b'\nW n\n' in stream:
        # Clipping path is not set
        return Image.new('1', image_size, 'white')
    matrix = stream.split(b'\n')[-1].split(b' ')
    if matrix[-1] != b'cm':
        # Something wrong
        return Image.new('1', image_size, 'white')
    matrix_width = float(matrix[0])
    zoom = width / matrix_width
    commands = stream.split(b'\nW n')[0].split(b'\n')
    surface = cairo.ImageSurface(cairo.FORMAT_A1, size[0], size[1])
    ctx = cairo.Context(surface)
    for command in commands:
        op = command.split(b' ')
        if op[-1] == b'm':
            x = float(op[0]) * zoom
            y = size[1] - float(op[1]) * zoom
            ctx.move_to(x, y)
        elif op[-1] == b'l':
            x = float(op[0]) * zoom
            y = size[1] - float(op[1]) * zoom
            ctx.line_to(x, y)
        elif op[-1] == b'c':
            x1 = float(op[0]) * zoom
            y1 = size[1] - float(op[1]) * zoom
            x2 = float(op[2]) * zoom
            y2 = size[1] - float(op[3]) * zoom
            x3 = float(op[4]) * zoom
            y3 = size[1] - float(op[5]) * zoom
            ctx.curve_to(x1, y1, x2, y2, x3, y3)
        elif op[-1] == b'v':
            x1, y1 = ctx.get_current_point()
            x2 = float(op[0]) * zoom
            y2 = size[1] - float(op[1]) * zoom
            x3 = float(op[2]) * zoom
            y3 = size[1] - float(op[3]) * zoom
            ctx.curve_to(x1, y1, x2, y2, x3, y3)
        elif op[-1] == b'y':
            x1 = float(op[0]) * zoom
            y1 = size[1] - float(op[1]) * zoom
            x3 = float(op[2]) * zoom
            y3 = size[1] - float(op[3]) * zoom
            ctx.curve_to(x1, y1, x3, y3, x3, y3)
        elif op[-1] == b're':
            x = float(op[0]) * zoom
            y = size[1] - float(op[1]) * zoom
            w = float(op[2]) * zoom
            h = float(op[3]) * zoom
            ctx.move_to(x, y)
            ctx.line_to(x + w, y)
            ctx.line_to(x + w, y - h)
            ctx.line_to(x, y - h)
            ctx.close_path()
        elif op[-1] == b'h':
            ctx.close_path()
    ctx.clip()
    ctx.rectangle(0, 0, size[0], size[1])
    ctx.set_source_rgb(1,1,1)
    ctx.fill()
    clipping_path =  Image.frombuffer('1', size, surface.get_data(), 'raw', '1;R' ,surface.get_stride())
    clipping_path = clipping_path.crop((image_pos[0], image_pos[1], image_pos[0] + image_size[0], image_pos[1] + image_size[1]))
    return clipping_path

def create_clipped_image_for_imagemask(imagemask, clipping_path):
    image_clipped = Image.new('LA', imagemask.size, (255, 0))
    image_clipped.paste(imagemask, mask=clipping_path)
    return image_clipped

def generate_image(config, doc, page, page_noimg, images, output_dir):
    zoom_list = []
    image_extract_list = []
    image_matrix_list = []
    image_rect_list = []
    img_xref_list = []
    image_type_list = []
    mode_list = []
    has_warning = False

    pagenum_str = str(page.number + 1).zfill(3)
    for image in images:
        img_xref = image[0]
        width = int(doc.xref_get_key(img_xref, "Width")[1])
        height = int(doc.xref_get_key(img_xref, "Height")[1])
        image_matrix = page.get_image_rects(img_xref, transform=True)[0][1]
        image_rect = page.get_image_rects(img_xref)[0]
        if image_matrix[1:3] != (0, 0):
            print(f'警告：{pagenum_str}-{img_xref}圖片旋轉或歪斜，輸出將與pdf不同')
            has_warning = True
        zoom = width / image_matrix[0]
        zoom_y = height / image_matrix[3]
        if zoom / zoom_y > 1.01 or zoom_y / zoom > 1.01:
            print(f'警告：{pagenum_str}-{img_xref}圖片寬高比改變')
            has_warning = True

        image_type, image_extract = extract_image(doc, img_xref, pagenum_str)
        if image_type == 'jpeg':
            if config['extract-jpeg']:
                with open(f"{output_dir}/{pagenum_str}-{img_xref}.jpg",'wb') as f:
                    f.write(image_extract)
            image_extract = Image.open(BytesIO(image_extract))
        elif image_type.startswith('mono'):
            image_extract = image_extract.convert('L')

        zoom_list.append(zoom)
        image_extract_list.append(image_extract)
        image_matrix_list.append(image_matrix)
        image_rect_list.append(image_rect)
        img_xref_list.append(img_xref)
        image_type_list.append(image_type)
        mode_list.append(image_extract.mode)

    zoom = zoom_list[find_largest_image(images)]
    for it in zoom_list:
        if math.ceil(page.rect[3] * zoom) != math.ceil(page.rect[3] * it):
            print(f'警告：第{pagenum_str}頁包含多張圖片，縮放程度不同')
            has_warning = True

    mode_merge = 'L'
    if 'RGB' in mode_list:
        mode_merge = 'RGB'
    elif 'CMYK' in mode_list:
        mode_merge = 'CMYK'

    rect_merge = page.rect
    if config['no-crop']:
        if len(images) > 1:
            print(f"警告：第{pagenum_str}頁包含多張圖片，使用'no-crop'選項可能導致圖片重疊")
            has_warning = True
        for image_rect in image_rect_list:
            if image_rect[0] < rect_merge[0]:
                rect_merge[0] = image_rect[0]
            if image_rect[1] < rect_merge[1]:
                rect_merge[1] = image_rect[1]
            if image_rect[2] > rect_merge[2]:
                rect_merge[2] = image_rect[2]
            if image_rect[3] > rect_merge[3]:
                rect_merge[3] = image_rect[3]

    if has_warning and config['render-image']:
        print(f"第{pagenum_str}頁使用渲染方式產生圖片")
        return render_image(page, zoom, colorspace=mode_merge, alpha=False)

    width_merge = math.ceil((rect_merge[2] - rect_merge[0]) * zoom)
    height_merge = math.ceil((rect_merge[3] - rect_merge[1]) * zoom)
    img_merge = Image.new(mode_merge, (width_merge, height_merge), 'white')
    for index in range(len(images)):
        image_pos = (round((image_matrix_list[index][4] - rect_merge[0]) * zoom), round((image_matrix_list[index][5] - rect_merge[1]) * zoom))
        if config['no-crop']:
            clipping_path = Image.new('1', image_extract_list[index].size, 'white')
        else:
            clipping_path = create_clipping_path_image(doc, page, images[index], (width_merge, height_merge), image_pos, image_extract_list[index].size)
        if image_type_list[index] == 'mono-mask':
            clipped_image = create_clipped_image_for_imagemask(image_extract_list[index], clipping_path)
            gray, alpha = clipped_image.split()
            invert = ImageOps.invert(gray)
            img_merge.paste(gray, image_pos, mask=invert)
        else:
            img_merge.paste(image_extract_list[index], image_pos, mask=clipping_path)
    img_noimg = render_image(page_noimg, zoom, colorspace=mode_merge)
    img_merge.paste(img_noimg, (int(-rect_merge[0] * zoom), int(-rect_merge[1] * zoom)), img_noimg)

    if all(image_type.startswith('mono') for image_type in image_type_list) and config['prefer-mono']:
        img_merge = img_merge.point(lambda i: i>127 and 255, mode='1')

    return img_merge

def save_pil_image(config, image, output_name):
    if config['save-png']:
        if image.mode == 'CMYK':
            image = image.convert('RGB')
        image.save(f"{output_name}.png")
    elif config['save-jxl']:
        if image.mode == '1':
            image = image.convert('L')
        if image.mode == 'CMYK':
            image = image.convert('RGB')
        image.save(f"{output_name}.jxl", lossless=True)
    elif config['save-tiff']:
        image.save(f"{output_name}.tiff", compression=config['save-tiff'])
    else:
        if image.mode == 'CMYK':
            image = image.convert('RGB')
        image.save(f"{output_name}.webp", lossless=True)

def convert_page(config, pagenum, output_dir):
    try:
        global doc
        page = doc[pagenum]
        images = page.get_images(full=True)
        if config['only-extract']:
            for image in images:
                save_extracted_image(config, doc, page, image, output_dir)
            return
        if not images:
            image = render_image(page, 600 / 72, alpha=False)
        else:
            global doc_noimg
            page_noimg = doc_noimg[pagenum]
            image = generate_image(config, doc, page, page_noimg, images, output_dir)
        save_pil_image(config, image, f"{output_dir}/{str(pagenum + 1).zfill(3)}")
    except Exception:
        print(traceback.format_exc())

def convert_page_init(file):
    try:
        global doc
        global doc_noimg
        doc = fitz.open(file)
        doc_noimg = fitz.open(file)
        for page in doc_noimg:
            remove_path_fill(doc_noimg, page)
            for image in page.get_images():
                xref = image[0]
                page.delete_image(xref)
        doc_noimg = fitz.open('pdf', doc_noimg.tobytes(garbage=1))
    except Exception:
        print(traceback.format_exc())

def main():
    if len(sys.argv) == 1:
        print('請選擇pdf檔')
        sys.exit(0)

    config = read_config()

    for file in sys.argv[1:]:
        if 'PDF2IMG_OUTPUT' in os.environ:
            output_dir = os.path.join(os.environ['PDF2IMG_OUTPUT'], file + "-img")
        else:
            output_dir = file + "-img"
        os.makedirs(output_dir, exist_ok=True)
        with fitz.open(file) as doc:
            page_count = doc.page_count
        with Pool(processes=config['processes'], initializer=convert_page_init, initargs=(file,)) as pool:
            pool.starmap(convert_page, zip(repeat(config), range(page_count), repeat(output_dir)))

main()
