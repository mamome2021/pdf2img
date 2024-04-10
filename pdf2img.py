#!/usr/bin/env python3
#第1版
import argparse
import sys
import os
import traceback
import fitz
from PIL import Image
from io import BytesIO
import math

parser = argparse.ArgumentParser()
parser.add_argument("input", type=str, nargs='+', help="pdf檔案名稱")
parser.add_argument("--tiff-compression", type=str, default="packbits", help="tiff壓縮方式")
parser.add_argument("--error", action='store_true', help="發生錯誤時中止程式")
args = parser.parse_args()

for file in args.input:
    doc = fitz.open(file)
    doc_noimg = fitz.open(file)
    for page in doc_noimg:
        for image in page.get_images():
            xref = image[0]
            page.delete_image(xref)
    doc_noimg = fitz.open('pdf', doc_noimg.tobytes())
    output_dir = file + "-img"
    os.makedirs(output_dir, exist_ok=True)

    for pagenum, page in enumerate(doc):
        page_noimg = doc_noimg[pagenum]
        if len(page.get_images()) > 1:
            print(f'警告：第{pagenum+1}頁包含多張圖片，輸出圖片只會包含一張圖片')
        for image in page.get_images():
            try:
                img_xref = image[0]
                width = int(doc.xref_get_key(img_xref, "Width")[1])
                height = int(doc.xref_get_key(img_xref, "Height")[1])
                cs_type = doc.xref_get_key(img_xref, "ColorSpace")[0]
                cs = doc.xref_get_key(img_xref, "ColorSpace")[1]
                output_name=f"{output_dir}/{pagenum+1}-{img_xref}"
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
                width_transform = image_matrix[0]
                zoom = width / width_transform
                pixmap_noimg = page_noimg.get_pixmap(matrix=fitz.Matrix(zoom,zoom), colorspace='GRAY', alpha=True)
                img_noimg = Image.frombytes('LA', [pixmap_noimg.width, pixmap_noimg.height], pixmap_noimg.samples)
                img_merge = Image.new(pil_image.mode, (math.ceil(page.rect[2] * zoom), math.ceil(page.rect[3] * zoom)),color='white')
                img_merge.paste(pil_image, (round(image_matrix[4]*zoom), round(image_matrix[5]*zoom)))
                img_merge.paste(img_noimg, (0, 0), img_noimg)
                img_merge.save(f"{output_name}.tiff", compression=args.tiff_compression)
            except Exception as e:
                print(traceback.format_exc())
                if args.error:
                    exit(1)
