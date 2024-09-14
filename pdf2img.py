#!/usr/bin/env python3
import tkinter
import tkinter.filedialog
from concurrent.futures import ProcessPoolExecutor
from concurrent.futures.process import BrokenProcessPool
from io import BytesIO
from itertools import repeat
import math
from multiprocessing import freeze_support
import os
import sys
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
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

def render_image(page, zoom, colorspace, alpha):
    if alpha == True and colorspace == 'CMYK':
        #CMYK with alpha channel is not supported by pillow
        colorspace = 'RGB'
    if colorspace == 'L':
        colorspace = 'GRAY'
    pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=colorspace, alpha=alpha)
    if colorspace == 'GRAY':
        colorspace = 'L'
    if not alpha:
        return Image.frombytes(colorspace, (pixmap.width, pixmap.height), pixmap.samples_mv)
    image = Image.frombytes(colorspace + 'a', (pixmap.width, pixmap.height), pixmap.samples_mv)
    image = image.convert(colorspace + 'A')
    return image

def get_image_colorspace(doc, img_xref):
    cs_type, cs = doc.xref_get_key(img_xref, "ColorSpace")
    if doc.xref_get_key(img_xref,"ImageMask")[1] == 'true':
        return '1'
    elif doc.xref_get_key(img_xref, "BitsPerComponent")[1] == '1':
        return '1'
    elif cs_type == 'xref':
        return 'RGB'
    elif cs == "/DeviceCMYK":
        return 'CMYK'
    elif cs == "/DeviceGray":
        return 'L'
    elif cs == "/DeviceRGB":
        return 'RGB'
    else:
        return 'RGB'

def extract_image(doc, img_xref, pagenum_str):
    width = int(doc.xref_get_key(img_xref, "Width")[1])
    height = int(doc.xref_get_key(img_xref, "Height")[1])
    cs_type = doc.xref_get_key(img_xref, "ColorSpace")[0]
    cs = doc.xref_get_key(img_xref, "ColorSpace")[1]
    if doc.xref_get_key(img_xref, "Filter")[1] == '/DCTDecode':
        if cs_type == 'xref':
            # Using xref_stream_raw directly produces image with inverted color
            # JOKER-我的同居小鬼(1) p3
            pixmap = fitz.Pixmap(doc, img_xref)
            pixmap = fitz.Pixmap(fitz.csRGB, pixmap)
            return "pil", Image.frombytes('RGB', (pixmap.width, pixmap.height), pixmap.samples_mv)
        if cs == "/DeviceCMYK":
            # Using xref_stream_raw directly produces image with inverted color
            pixmap = fitz.Pixmap(doc, img_xref)
            return "pil", Image.frombytes('CMYK', (pixmap.width, pixmap.height), pixmap.samples_mv)
        return "jpeg", doc.xref_stream_raw(img_xref)
    elif doc.xref_get_key(img_xref,"ImageMask")[1] == 'true':
        return "mask", Image.frombytes('1', (width, height), doc.xref_stream(img_xref))
    elif doc.xref_get_key(img_xref, "BitsPerComponent")[1] == '1':
        return "pil", Image.frombytes('1', (width, height), doc.xref_stream(img_xref))
    elif cs_type == 'xref':
        print(f"警告：{pagenum_str}-{img_xref} xref cs")
        # 太難了不會做，用第一版的方法
        img_dict = doc.extract_image(img_xref)
        img_data = img_dict["image"]
        return "pil", Image.open(BytesIO(img_data))
    elif cs == "/DeviceCMYK":
        return "pil", Image.frombytes('CMYK', (width, height), doc.xref_stream(img_xref))
    elif cs == "/DeviceGray":
        return "pil", Image.frombytes('L', (width, height), doc.xref_stream(img_xref))
    elif cs == "/DeviceRGB":
        return "pil", Image.frombytes('RGB', (width, height), doc.xref_stream(img_xref))
    else:
        print(f"警告：{pagenum_str}-{img_xref}未知色彩空間", cs)
        # 其他，還沒做，用第一版的方法
        img_dict = doc.extract_image(img_xref)
        img_data = img_dict["image"]
        return "pil", Image.open(BytesIO(img_data))

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
    clipping_path = Image.frombuffer('1', size, surface.get_data(), 'raw', '1;R' ,surface.get_stride())
    clipping_path = clipping_path.crop((image_pos[0], image_pos[1], image_pos[0] + image_size[0], image_pos[1] + image_size[1]))
    return clipping_path

def create_clipped_image_for_imagemask(imagemask, clipping_path):
    image_clipped = Image.new('1', imagemask.size, 255)
    image_clipped.paste(imagemask, mask=clipping_path)
    return image_clipped

def generate_image(config, doc, page, page_noimg, images, output_dir):
    zoom_list = []
    image_matrix_list = []
    image_rect_list = []
    has_warning = False

    is_mono = True
    mode_merge = 'L'
    pagenum_str = str(page.number + 1).zfill(3)
    for image in images:
        img_xref = image[0]
        img_name = image[7]
        width = int(doc.xref_get_key(img_xref, "Width")[1])
        height = int(doc.xref_get_key(img_xref, "Height")[1])
        image_rect, image_matrix = page.get_image_bbox(img_name, transform=True)
        if image_matrix[1:3] != (0, 0):
            print(f'警告：{pagenum_str}-{img_xref}圖片旋轉或歪斜，輸出將與pdf不同')
            has_warning = True
        zoom = width / image_matrix[0]
        zoom_y = height / image_matrix[3]
        if zoom / zoom_y > 1.01 or zoom_y / zoom > 1.01:
            print(f'警告：{pagenum_str}-{img_xref}圖片寬高比改變')
            has_warning = True

        image_colorspace = get_image_colorspace(doc, img_xref)
        if image_colorspace != '1':
            is_mono = False
        if mode_merge =='L' and image_colorspace == 'CMYK':
            mode_merge = 'CMYK'
        if image_colorspace == 'RGB':
            mode_merge = 'RGB'

        zoom_list.append(zoom)
        image_matrix_list.append(image_matrix)
        image_rect_list.append(image_rect)

    zoom = zoom_list[find_largest_image(images)]
    for it in zoom_list:
        if math.ceil(page.rect[3] * zoom) != math.ceil(page.rect[3] * it):
            print(f'警告：第{pagenum_str}頁包含多張圖片，縮放程度不同')
            has_warning = True

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
        img_xref = images[index][0]
        image_type, image_extract = extract_image(doc, img_xref, pagenum_str)
        if image_type == 'jpeg':
            if config['extract-jpeg']:
                with open(f"{output_dir}/{pagenum_str}-{img_xref}.jpg",'wb') as f:
                    f.write(image_extract)
            image_extract = Image.open(BytesIO(image_extract))
        image_pos = (round((image_matrix_list[index][4] - rect_merge[0]) * zoom), round((image_matrix_list[index][5] - rect_merge[1]) * zoom))
        if config['no-crop']:
            clipping_path = Image.new('1', image_extract.size, 'white')
        else:
            clipping_path = create_clipping_path_image(doc, page, images[index], (width_merge, height_merge), image_pos, image_extract.size)
        if image_type == 'mask':
            clipped_image = create_clipped_image_for_imagemask(image_extract, clipping_path)
            del image_extract
            del clipping_path
            invert = ImageOps.invert(clipped_image)
            img_merge.paste(clipped_image, image_pos, mask=invert)
            del clipped_image
            del invert
        else:
            img_merge.paste(image_extract, image_pos, mask=clipping_path)
            del image_extract
            del clipping_path
    img_noimg = render_image(page_noimg, zoom, colorspace=mode_merge, alpha=True)
    img_merge.paste(img_noimg, (int(-rect_merge[0] * zoom), int(-rect_merge[1] * zoom)), img_noimg)
    del img_noimg

    if is_mono and config['prefer-mono']:
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
        if not doc:
            # Failed in convert_page_init()
            return 0
        page = doc[pagenum]
        images = page.get_images(full=True)
        if config['only-extract']:
            for image in images:
                save_extracted_image(config, doc, page, image, output_dir)
            return 1
        if not images:
            image = render_image(page, 600 / 72, colorspace='GRAY', alpha=False)
        else:
            global doc_noimg
            page_noimg = doc_noimg[pagenum]
            image = generate_image(config, doc, page, page_noimg, images, output_dir)
        save_pil_image(config, image, f"{output_dir}/{str(pagenum + 1).zfill(3)}")
        return 1
    except Exception:
        print(traceback.format_exc())
        return 0

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
        del doc

def gui(config):
    def open_pdf_file():
        name = tkinter.filedialog.askopenfilename()
        if name:
            pdf_file.delete('1.0', 'end')
            pdf_file.insert('end', name)

    def open_output_dir():
        name = tkinter.filedialog.askdirectory()
        if name:
            output_dir_text.delete('1.0', 'end')
            output_dir_text.insert('end', name)

    def convert():
        file = pdf_file.get('1.0' ,'end-1c')
        if not file:
            return
        try:
            with fitz.open(file) as doc:
                page_count = doc.page_count
        except fitz.FileNotFoundError:
            messagebox.showinfo(message='找不到檔案')
            return
        output_dir = output_dir_text.get('1.0', 'end-1c')
        if not output_dir:
            output_dir = file + "-img"
        os.makedirs(output_dir, exist_ok=True)
        config['processes'] = processes.get()
        config['only-extract'] = only_extract.get()
        config['render-image'] = render_image.get()
        config['no-crop'] = no_crop.get()
        config['extract-jpeg'] = extract_jpeg.get()
        config['prefer-mono'] = prefer_mono.get()
        config['save-jxl'] = save_jxl.get()
        config['save-png'] = save_png.get()
        config['save-tiff'] = save_tiff.get('1.0' ,'end-1c')
        failed_page = []
        # Use ProcessPoolExecutor instead of multiprocessing.Pool
        # to detect error of process killed due to low memory
        with ProcessPoolExecutor(max_workers=config['processes'], initializer=convert_page_init,
                                 initargs=(file,)) as pool:
            try:
                for idx, result in enumerate(pool.map(convert_page, repeat(config), range(page_count), repeat(output_dir))):
                    if result != 1:
                        failed_page.append(str(idx + 1))
            except BrokenProcessPool:
                messagebox.showinfo(message='BrokenProcessPool: 可能記憶體不足')
                return

        message = '轉換完成'
        if failed_page:
            message += f"，第{', '.join(failed_page)}頁轉換失敗"
        messagebox.showinfo(message=message)

    root = Tk()
    root.title('pdf2img')

    processes = IntVar(value=config['processes'])
    only_extract = BooleanVar(value=config['only-extract'])
    render_image = BooleanVar(value=config['render-image'])
    no_crop = BooleanVar(value=config['no-crop'])
    extract_jpeg = BooleanVar(value=config['extract-jpeg'])
    prefer_mono = BooleanVar(value=config['prefer-mono'])
    save_jxl = BooleanVar(value=config['save-jxl'])
    save_png = BooleanVar(value=config['save-png'])


    ttk.Button(root, text="要轉換的PDF檔", command=open_pdf_file).grid(column=0, row=0)
    pdf_file = Text(root, height=1)
    pdf_file.grid(column=1, row=0)
    ttk.Button(root, text="輸出資料夾", command=open_output_dir).grid(column=0, row=1)
    output_dir_text = Text(root, height=1)
    output_dir_text.grid(column=1, row=1)
    ttk.Label(root, text='進程數（請注意記憶體是否足夠）').grid(column=0, row=2)
    ttk.Spinbox(from_=1, to=8, textvariable=processes).grid(column=1, row=2)
    ttk.Label(root, text='只提取原圖，不疊加渲染圖').grid(column=0, row=3)
    ttk.Checkbutton(root,variable=only_extract).grid(column=1, row=3)
    ttk.Label(root, text='如果無法完美提取，則使用渲染方式產生圖片').grid(column=0, row=4)
    ttk.Checkbutton(root,variable=render_image).grid(column=1, row=4)
    ttk.Label(root, text='不裁切超出pdf頁面的原圖，並忽略clipping path').grid(column=0, row=5)
    ttk.Checkbutton(root,variable=no_crop).grid(column=1, row=5)
    ttk.Label(root, text='若圖片為jpeg，也提取出未疊加渲染圖的原圖').grid(column=0, row=6)
    ttk.Checkbutton(root,variable=extract_jpeg).grid(column=1, row=6)
    ttk.Label(root, text='若原圖為位圖，以位圖格式儲存').grid(column=0, row=7)
    ttk.Checkbutton(root,variable=prefer_mono).grid(column=1, row=7)
    ttk.Label(root, text='以JPEG XL格式儲存').grid(column=0, row=8)
    ttk.Checkbutton(root,variable=save_jxl).grid(column=1, row=8)
    ttk.Label(root, text='以png格式儲存').grid(column=0, row=9)
    ttk.Checkbutton(root,variable=save_png).grid(column=1, row=9)
    ttk.Label(root, text='以tiff格式儲存，並指定壓縮方式').grid(column=0, row=10)
    save_tiff = Text(root, height=1)
    save_tiff.grid(column=1, row=10)
    save_tiff.insert('end-1c',config['save-tiff'])
    button_convert = ttk.Button(root, text="轉換", command=convert)
    button_convert.grid(column=1, row=11)
    root.mainloop()

def main():
    config = read_config()

    if len(sys.argv) == 1:
        gui(config)
        sys.exit(0)

    for file in sys.argv[1:]:
        if 'PDF2IMG_OUTPUT' in os.environ:
            output_dir = os.path.join(os.environ['PDF2IMG_OUTPUT'], file + "-img")
        else:
            output_dir = file + "-img"
        os.makedirs(output_dir, exist_ok=True)
        with fitz.open(file) as doc:
            page_count = doc.page_count
        # Use ProcessPoolExecutor instead of multiprocessing.Pool
        # to detect error of process killed due to low memory
        with ProcessPoolExecutor(max_workers=config['processes'], initializer=convert_page_init, initargs=(file,)) as pool:
            for idx, result in enumerate(pool.map(convert_page, repeat(config), range(page_count), repeat(output_dir))):
                if result != 1:
                    print(f'第{idx + 1}頁轉換失敗')

if __name__ == '__main__':
    freeze_support()
    main()
