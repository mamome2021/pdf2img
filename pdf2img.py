#!/usr/bin/env python3
import sys
import os
import traceback
import fitz
from PIL import Image
from io import BytesIO
import math

def read_config():
    config = {'error': False,
              'only-extract': False,
              'single-image': False,
              'no-crop': False,
              'remove-path-fill': False,
              'extract-jpeg': False,
              'small-output': False,
              'prefer-mono': False,
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
            elif option[0] == 'only-extract':
                config['only-extract'] = True
            elif option[0] == 'single-image':
                config['single-image'] = True
            elif option[0] == 'no-crop':
                config['no-crop'] = True
            elif option[0] == 'remove-path-fill':
                config['remove-path-fill'] = True
            elif option[0] == 'extract-jpeg':
                config['extract-jpeg'] = True
            elif option[0] == 'small-output':
                config['small-output'] = True
            elif option[0] == 'prefer-mono':
                config['prefer-mono'] = True
            elif option[0] == 'prefer-png':
                config['prefer-png'] = True
            elif option[0] == 'tiff-compression':
                config['tiff-compression'] = option[1]
    except FileNotFoundError:
        pass
    except Exception:
        print(traceback.format_exc())
    return config

def remove_path_fill(doc, page):
    images = page.get_images(full=True)
    if len(images) == 0:
        return
    image = images[0]
    image_name = image[7]
    ref = image[9]
    if ref == 0:
        # Image is directly referenced by the page
        ref = page.get_contents()[0]
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
    return images[index]

def render_image(config, page, zoom, colorspace='GRAY', alpha=True):
    pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=colorspace, alpha=alpha)
    if not alpha:
        return Image.frombytes('L', [pixmap.width, pixmap.height], pixmap.samples)
    image = Image.frombytes('La', [pixmap.width, pixmap.height], pixmap.samples)
    image = image.convert('LA')
    return image

def extract_image(doc, img_xref, output_name):
    width = int(doc.xref_get_key(img_xref, "Width")[1])
    height = int(doc.xref_get_key(img_xref, "Height")[1])
    cs_type = doc.xref_get_key(img_xref, "ColorSpace")[0]
    cs = doc.xref_get_key(img_xref, "ColorSpace")[1]
    if doc.xref_get_key(img_xref, "Filter")[1] == '/DCTDecode':
        if cs == "/DeviceCMYK":
            print(output_name, "jpeg-cmyk")
            pixmap = fitz.Pixmap(doc, img_xref)
            return "cmyk", Image.frombytes('CMYK', (pixmap.width, pixmap.height), pixmap.samples)
        else:
            print(output_name, "jpeg")
            return "jpeg", doc.xref_stream_raw(img_xref)
    elif doc.xref_get_key(img_xref,"ImageMask")[1] == 'true' or doc.xref_get_key(img_xref, "BitsPerComponent")[1] == '1':
        print(output_name, "mono")
        return "mono", Image.frombytes('1', (width, height), doc.xref_stream(img_xref))
    elif cs_type == 'xref':
        print(output_name, "xref cs")
        # 太難了不會做，用第一版的方法
        img_dict = doc.extract_image(img_xref)
        img_data = img_dict["image"]
        return "rgb", Image.open(BytesIO(img_data))
    elif cs == "/DeviceCMYK":
        print(output_name, "cmyk")
        return "cmyk", Image.frombytes('CMYK', (width, height), doc.xref_stream(img_xref))
    elif cs == "/DeviceGray":
        print(output_name, "gray")
        return "gray", Image.frombytes('L', (width, height), doc.xref_stream(img_xref))
    elif cs == "/DeviceRGB":
        print(output_name,"rgb")
        return "rgb", Image.frombytes('RGB', (width, height), doc.xref_stream(img_xref))
    else:
        print(output_name,"wip:", cs)
        # 其他，還沒做，用第一版的方法
        img_dict = doc.extract_image(img_xref)
        img_data = img_dict["image"]
        return "rgb", Image.open(BytesIO(img_data))

def save_extracted_image(config, doc, page, image, output_dir):
    img_xref = image[0]
    output_name = f"{output_dir}/{page.number+1}-{img_xref}"
    image_type, image_extract = extract_image(doc, img_xref, output_name)
    if image_type == 'jpeg':
        with open(f"{output_name}.jpg",'wb') as f:
            f.write(image_extract)
    else:
        save_pil_image(config, image_extract, output_name)

def generate_image(config, doc, page, page_noimg, image, output_dir):
    img_xref = image[0]
    width = int(doc.xref_get_key(img_xref, "Width")[1])
    height = int(doc.xref_get_key(img_xref, "Height")[1])
    output_name = f"{output_dir}/{page.number+1}-{img_xref}"
    image_matrix = page.get_image_rects(img_xref, transform=True)[0][1]
    if image_matrix[1:3] != (0, 0):
        print(output_name, '警告：圖片旋轉或歪斜，輸出將與pdf不同')
    zoom = width / image_matrix[0]
    zoom_y = height / image_matrix[3]
    if zoom / zoom_y > 1.01 or zoom_y / zoom > 1.01:
        print('警告：圖片寬高比改變')
    img_noimg = render_image(config, page_noimg, zoom)

    image_type, image_extract = extract_image(doc, img_xref, output_name)
    if image_type == 'jpeg':
        image_extract = Image.open(BytesIO(image_extract))
        if config['extract-jpeg']:
            open(f"{output_name}.jpg",'wb').write(image_extract)
    elif image_type == 'mono':
        image_extract = image_extract.convert('L')

    if not config['no-crop']:
        img_merge = Image.new(image_extract.mode, (math.ceil(page.rect[2] * zoom), math.ceil(page.rect[3] * zoom)), color='white')
        img_merge.paste(image_extract, (round(image_matrix[4] * zoom), round(image_matrix[5] * zoom)))
        img_merge.paste(img_noimg, (0, 0), img_noimg)
    else:
        image_rect = page.get_image_rects(img_xref)[0]
        width_merge = max(page.rect[2], image_rect[2]) - min(page.rect[0], image_rect[0])
        height_merge = max(page.rect[3], image_rect[3]) - min(page.rect[1], image_rect[1])
        x_offset = min(image_rect[0], 0)
        y_offset = min(image_rect[1], 0)
        img_merge = Image.new(image_extract.mode, (math.ceil(width_merge * zoom), math.ceil(height_merge * zoom)), color='white')
        img_merge.paste(image_extract, (round(max(image_matrix[4], 0) * zoom), round(max(image_matrix[5], 0) * zoom)))
        img_merge.paste(img_noimg, (round(-x_offset * zoom), round(-y_offset * zoom)), img_noimg)
    if image_type == 'mono' and config['prefer-mono']:
        img_merge = img_merge.point(lambda i: i>127 and 255, mode='1')

    return img_merge

def save_pil_image(config, image, output_name):
    if config['small-output']:
        if image.mode == 'CMYK':
            image.save(f"{output_name}.tiff", compression='tiff_lzw')
        elif image.mode == '1':
            io_group4=BytesIO()
            io_png=BytesIO()
            io_lzw=BytesIO()
            image.save(io_group4, format='tiff', compression='group4')
            image.save(io_png, format='png')
            image.save(io_lzw, format='tiff', compression='tiff_lzw')
            if io_lzw.getbuffer().nbytes < io_png.getbuffer().nbytes and io_lzw.getbuffer().nbytes < io_group4.getbuffer().nbytes:
                image.save(f"{output_name}.tiff", compression='tiff_lzw')
            elif io_group4.getbuffer().nbytes < io_png.getbuffer().nbytes:
                image.save(f"{output_name}.tiff", compression='group4')
            else:
                image.save(f"{output_name}.png")
        else:
            io_lzw=BytesIO()
            io_png=BytesIO()
            image.save(io_lzw, format='tiff', compression='tiff_lzw')
            image.save(io_png, format='png')
            if io_lzw.getbuffer().nbytes < io_png.getbuffer().nbytes:
                image.save(f"{output_name}.tiff", compression='tiff_lzw')
            else:
                image.save(f"{output_name}.png")
    else:
        if config['prefer-png'] == True and image.mode != 'CMYK':
            image.save(f"{output_name}.png")
        else:
            image.save(f"{output_name}.tiff", compression=config['tiff-compression'])

def main():
    if len(sys.argv) == 1:
        print('請選擇pdf檔')
    
    config = read_config()

    for file in sys.argv[1:]:
        doc = fitz.open(file)
        doc_noimg = fitz.open(file)
        for page in doc_noimg:
            if config['remove-path-fill']:
                remove_path_fill(doc_noimg, page)
            for image in page.get_images():
                xref = image[0]
                page.delete_image(xref)
        doc_noimg = fitz.open('pdf', doc_noimg.tobytes(garbage=1))
        output_dir = file + "-img"
        os.makedirs(output_dir, exist_ok=True)

        for pagenum, page in enumerate(doc):
            try:
                page_noimg = doc_noimg[pagenum]
                images = page.get_images()
                pagenum_str = str(pagenum + 1).zfill(3)

                if config['only-extract']:
                    for image in images:
                        save_extracted_image(config, doc, page, image, output_dir)
                    continue

                if not images:
                    print(f'警告：第{pagenum+1}頁沒有圖片，使用600dpi渲染')
                    image = render_image(config, page, 600 / 72, alpha=False)
                    save_pil_image(config, image, f"{output_dir}/{pagenum_str}")
                    continue
                if len(images) > 1:
                    print(f'警告：第{pagenum+1}頁包含多張圖片，輸出圖片只會包含一張圖片')
                    if config['single-image']:
                        images = [find_largest_image(images)]
                for image in images:
                    img_generated = generate_image(config, doc, page, page_noimg, image, output_dir)
                    if len(images) == 1:
                        output_name = f"{output_dir}/{pagenum_str}"
                    else:
                        output_name = f"{output_dir}/{pagenum_str}-{image[0]}"
                    save_pil_image(config, img_generated, output_name)
            except Exception as e:
                print(traceback.format_exc())
                if config['error']:
                    exit(1)

main()
