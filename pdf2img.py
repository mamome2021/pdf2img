#!/usr/bin/env python3
import sys
import os
import traceback
import fitz
from PIL import Image, ImageOps
from io import BytesIO
import math
import cairo

def read_config():
    config = {'error': False,
              'only-extract': False,
              'no-crop': False,
              'remove-path-fill': False,
              'extract-jpeg': False,
              'small-output': False,
              'prefer-mono': False,
              'prefer-png': False,
              'tiff-compression': 'packbits'}
    try:
        if 'PDF2IMG_CONFIG' in os.environ:
            config_filename = os.environ['PDF2IMG_CONFIG']
        else:
            config_filename = os.path.abspath(os.path.join(os.path.dirname(__file__), 'config-pdf2img.txt'))
        with open(config_filename, 'r', encoding='utf-8') as config_file:
            lines = config_file.read().split('\n')
        for line in lines:
            option = line.split()
            if len(option) == 0:
                continue
            if option[0] == 'error':
                config['error'] = True
            elif option[0] == 'only-extract':
                config['only-extract'] = True
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
    return index

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
    elif doc.xref_get_key(img_xref,"ImageMask")[1] == 'true':
        print(output_name, "mask")
        return "mono-mask", Image.frombytes('1', (width, height), doc.xref_stream(img_xref))
    elif doc.xref_get_key(img_xref, "BitsPerComponent")[1] == '1':
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
    pagenum_str = str(page.number + 1).zfill(3)
    output_name = f"{output_dir}/{pagenum_str}-{img_xref}"
    image_type, image_extract = extract_image(doc, img_xref, output_name)
    if image_type == 'jpeg':
        with open(f"{output_name}.jpg",'wb') as f:
            f.write(image_extract)
    else:
        save_pil_image(config, image_extract, output_name)

def apply_clipping_path(doc, image, image_extract):
    image_name = image[7]
    image_xref = image[0]
    width = int(doc.xref_get_key(image_xref, "Width")[1])
    referencer = image[9]
    stream = doc.xref_stream(referencer).split(f'\n/{image_name} Do\n'.encode())[0].split(b'\nQ\n')[-1]
    if not b'\nW n\n' in stream:
        # Clipping path is not set, return original image
        return image_extract
    matrix = stream.split(b'\n')[-1].split(b' ')
    if matrix[-1] != b'cm':
        # Something wrong, return original image
        return image_extract
    matrix_width = float(matrix[0])
    zoom = width / matrix_width
    commands = stream.split(b'\nW n')[0].split(b'\n')
    surface = cairo.ImageSurface(cairo.FORMAT_A1, image_extract.width, image_extract.height)
    ctx = cairo.Context (surface)
    for command in commands:
        op = command.split(b' ')
        if op[-1] == b'm':
            x = float(op[0]) * zoom
            y = image_extract.height - float(op[1]) * zoom
            ctx.move_to(x, y)
        elif op[-1] == b'l':
            x = float(op[0]) * zoom
            y = image_extract.height - float(op[1]) * zoom
            ctx.line_to(x, y)
        elif op[-1] == b'c':
            x1 = float(op[0]) * zoom
            y1 = image_extract.height - float(op[1]) * zoom
            x2 = float(op[2]) * zoom
            y2 = image_extract.height - float(op[3]) * zoom
            x3 = float(op[4]) * zoom
            y3 = image_extract.height - float(op[5]) * zoom
            ctx.curve_to(x1, y1, x2, y2, x3, y3)
        elif op[-1] == b'v':
            x1, y1 = ctx.get_current_point()
            x2 = float(op[0]) * zoom
            y2 = image_extract.height - float(op[1]) * zoom
            x3 = float(op[2]) * zoom
            y3 = image_extract.height - float(op[3]) * zoom
            ctx.curve_to(x1, y1, x2, y2, x3, y3)
        elif op[-1] == b'y':
            x1 = float(op[0]) * zoom
            y1 = image_extract.height - float(op[1]) * zoom
            x3 = float(op[2]) * zoom
            y3 = image_extract.height - float(op[3]) * zoom
            ctx.curve_to(x1, y1, x3, y3, x3, y3)
        elif op[-1] == b're':
            x = float(op[0]) * zoom
            y = image_extract.height - float(op[1]) * zoom
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
    ctx.rectangle(0, 0, image_extract.width, image_extract.height)
    ctx.set_source_rgb(1,1,1)
    ctx.fill()
    clip_pil = Image.frombuffer('1', image_extract.size, surface.get_data(), 'raw', '1;R' ,surface.get_stride())
    image_clipped = Image.new(image_extract.mode, image_extract.size, 'white')
    image_clipped.putalpha(0)
    image_clipped.paste(image_extract, mask=clip_pil)
    return image_clipped

def generate_image(config, doc, page, page_noimg, images, output_dir):
    zoom_list = []
    image_extract_list = []
    image_matrix_list = []
    img_xref_list = []
    image_type_list = []
    
    for image in images:
        img_xref = image[0]
        width = int(doc.xref_get_key(img_xref, "Width")[1])
        height = int(doc.xref_get_key(img_xref, "Height")[1])
        pagenum_str = str(page.number + 1).zfill(3)
        output_name = f"{output_dir}/{pagenum_str}-{img_xref}"
        image_matrix = page.get_image_rects(img_xref, transform=True)[0][1]
        if image_matrix[1:3] != (0, 0):
            print(output_name, '警告：圖片旋轉或歪斜，輸出將與pdf不同')
        zoom = width / image_matrix[0]
        zoom_y = height / image_matrix[3]
        if zoom / zoom_y > 1.01 or zoom_y / zoom > 1.01:
            print('警告：圖片寬高比改變')

        image_type, image_extract = extract_image(doc, img_xref, output_name)
        if image_type == 'jpeg':
            if config['extract-jpeg']:
                with open(f"{output_name}.jpg",'wb') as f:
                    f.write(image_extract)
            image_extract = Image.open(BytesIO(image_extract))
        elif image_type.startswith('mono'):
            image_extract = image_extract.convert('L')

        zoom_list.append(zoom)
        image_extract_list.append(image_extract)
        image_matrix_list.append(image_matrix)
        img_xref_list.append(img_xref)
        image_type_list.append(image_type)

    zoom = zoom_list[find_largest_image(images)]
    for it in zoom_list:
        if math.ceil(page.rect[3] * zoom) != math.ceil(page.rect[3] * it):
            print(f'警告：第{pagenum_str}頁包含多張圖片，縮放程度不同')

    img_noimg = render_image(config, page_noimg, zoom)
    if not config['no-crop']:
        # TODO: find good mode
        width_merge = math.ceil(page.rect[2] * zoom)
        height_merge = math.ceil(page.rect[3] * zoom)
        img_merge = Image.new(image_extract_list[0].mode, (width_merge, height_merge), 'white')
        for index in range(len(images)):
            # TODO: putalpha converts CMYK to RGBA
            image_clip = Image.new(image_extract_list[index].mode, (width_merge, height_merge), 'white')
            image_clip.putalpha(0)
            image_clip.paste(image_extract_list[index], (round(image_matrix_list[index][4] * zoom), round(image_matrix_list[index][5] * zoom)))
            image_clip = apply_clipping_path(doc, images[index], image_clip)
            if image_type_list[index] == 'mono-mask':
                gray, alpha = image_clip.split()
                invert = ImageOps.invert(gray)
                img_merge.paste(gray, mask=invert)
            else:
                img_merge.paste(image_clip)
        img_merge.paste(img_noimg, (0, 0), img_noimg)

    else:
        # TODO: make no-crop work
        # TODO: make image_rect a list
        image_rect = page.get_image_rects(img_xref_list[0])[0]
        width_merge = max(page.rect[2], image_rect[2]) - min(page.rect[0], image_rect[0])
        height_merge = max(page.rect[3], image_rect[3]) - min(page.rect[1], image_rect[1])
        x_offset = min(image_rect[0], 0)
        y_offset = min(image_rect[1], 0)
        # TODO: find good mode
        img_merge = Image.new(image_extract_list[0].mode, (math.ceil(width_merge * zoom), math.ceil(height_merge * zoom)), color='white')
        for index in range(len(images)):
            img_merge.paste(image_extract_list[index], (round(max(image_matrix_list[index][4], 0) * zoom), round(max(image_matrix_list[index][5], 0) * zoom)))
        img_merge.paste(img_noimg, (round(-x_offset * zoom), round(-y_offset * zoom)), img_noimg)
    if all(image_type.startswith('mono') for image_type in image_type_list) and config['prefer-mono']:
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
        if 'PDF2IMG_OUTPUT' in os.environ:
            output_dir = os.environ['PDF2IMG_OUTPUT']
        else:
            output_dir = file + "-img"
        os.makedirs(output_dir, exist_ok=True)

        for pagenum, page in enumerate(doc):
            try:
                page_noimg = doc_noimg[pagenum]
                images = page.get_images(full=True)
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
                img_generated = generate_image(config, doc, page, page_noimg, images, output_dir)
                save_pil_image(config, img_generated, f"{output_dir}/{pagenum_str}")
            except Exception as e:
                print(traceback.format_exc())
                if config['error']:
                    exit(1)

main()
