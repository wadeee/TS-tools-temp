import hashlib
import os
import shutil
import zipfile
from uuid import uuid1

import cv2
import pythoncom
from fitz import fitz
from flask import Flask, render_template, request, send_file
from skimage import io
from win32com.client import Dispatch

app = Flask(__name__)


@app.route('/')
def index():
    return render_template("index.html")


@app.route('/a', methods=['POST'])
def a():
    a4width = 21
    a4height = 29.7
    a4margintopbot = 2.54
    a4marginleftright = 3.18
    mypath = getmypath()
    if not os.path.exists(mypath):
        os.makedirs(mypath)

    docname = request.files.get('originFileA').filename
    docpathname = os.path.join(mypath, docname)
    namepre = os.path.splitext(docname)[0]
    pdfname = namepre + ".pdf"
    pdfpathname = os.path.join(mypath, pdfname)
    imagepath = os.path.join(mypath, namepre)
    zipfilename = namepre + '.zip'
    zippath = os.path.join(mypath, 'zipfiles')
    zippathname = os.path.join(zippath, zipfilename)

    request.files.get('originFileA').save(docpathname)

    pythoncom.CoInitialize()
    word = Dispatch('Word.Application')
    worddoc = word.Documents.Open(docpathname, ReadOnly=1)
    worddoc.SaveAs(pdfpathname, FileFormat=17)
    worddoc.Close()
    os.remove(docpathname)

    if os.path.exists(zippath):
        shutil.rmtree(zippath)
    os.makedirs(zippath)
    if not os.path.exists(imagepath):
        os.makedirs(imagepath)

    filedoc = fitz.open(pdfpathname)
    for pg in range(filedoc.pageCount):
        page = filedoc[pg]
        rotate = int(0)
        zoom_x = 2
        zoom_y = 2
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        rect = page.rect
        cliptl = rect.tr * a4marginleftright / a4width + rect.bl * a4margintopbot / a4height
        clipbr = rect.tr * (a4width - a4marginleftright) / a4width + rect.bl * (a4height - a4margintopbot) / a4height
        clip = fitz.Rect(cliptl, clipbr)
        pix = page.getPixmap(matrix=mat, alpha=False, clip=clip)

        pix.writePNG(os.path.join(imagepath, '{}.png'.format(pg)))
    filedoc.close()

    zipimages(zippathname, imagepath)
    shutil.rmtree(imagepath)
    os.remove(pdfpathname)

    return send_file(zippathname,
                     mimetype='application/x-zip-compressed;charset=utf-8',
                     attachment_filename=zipfilename,
                     as_attachment=True)


@app.route('/b', methods=['POST'])
def b():
    mypath = getmypath()
    if not os.path.exists(mypath):
        os.makedirs(mypath)
    pdfname = request.files.get('originFileB').filename
    pdfpathname = os.path.join(mypath, pdfname)
    namepre = os.path.splitext(pdfname)[0]

    imagepath = os.path.join(mypath, namepre)
    zipfilename = namepre + '.zip'
    zippath = os.path.join(mypath, 'zipfiles')
    zippathname = os.path.join(zippath, zipfilename)

    request.files.get('originFileB').save(pdfpathname)

    if os.path.exists(zippath):
        shutil.rmtree(zippath)
    os.makedirs(zippath)
    if not os.path.exists(imagepath):
        os.makedirs(imagepath)

    filedoc = fitz.open(pdfpathname)
    for pg in range(filedoc.pageCount):
        page = filedoc[pg]
        rotate = int(0)
        zoom_x = 2
        zoom_y = 2
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        rect = page.rect
        clip = fitz.Rect(rect.tl, rect.br)
        pix = page.getPixmap(matrix=mat, alpha=False, clip=clip)

        pix.writePNG(os.path.join(imagepath, '{}.jpg'.format(pg)))
    filedoc.close()

    zipimages(zippathname, imagepath)
    shutil.rmtree(imagepath)
    os.remove(pdfpathname)

    return send_file(zippathname,
                     mimetype='application/x-zip-compressed;charset=utf-8',
                     attachment_filename=zipfilename,
                     as_attachment=True)


@app.route('/c', methods=['POST'])
def c():
    images = request.files.getlist('originFilesC')
    mypath = getmypath()
    imagepath = os.path.join(mypath, 'images')
    if not os.path.exists(imagepath):
        os.makedirs(imagepath)
    for image in images:
        img_re = corpmargin(io.imread(image))
        io.imsave(os.path.join(imagepath, image.filename), img_re)

    zippathname = os.path.join(mypath, 'result.zip')
    zipimages(zippathname, imagepath)
    return send_file(zippathname,
                     mimetype='application/x-zip-compressed;charset=utf-8',
                     attachment_filename='result.zip',
                     as_attachment=True)


@app.route('/d', methods=['POST'])
def d():
    images = request.files.getlist('originFilesD')
    mypath = getmypath()
    imagepath = os.path.join(mypath, 'images')
    if not os.path.exists(imagepath):
        os.makedirs(imagepath)

    readimages = []
    readimagesresult = []
    width = 0
    for image in images:
        readimages.append(io.imread(image))

    for readimage in readimages:
        (x, y, z) = readimage.shape
        print(x, y)
        width = max(width, y)

    for readimage in readimages:
        (x, y, z) = readimage.shape
        if y == width:
            readimagesresult.append(readimage)
        else:
            readimagesresult.append(cv2.resize(readimage, (width, round(x * width / y)), interpolation=cv2.INTER_AREA))

    resultimage = cv2.vconcat(readimagesresult)

    resultimagepathname = os.path.join(imagepath, 'resultimage.jpg')
    io.imsave(resultimagepathname, resultimage)
    return send_file(resultimagepathname,
                     mimetype='image/jpeg',
                     attachment_filename='resultimage.jpg',
                     as_attachment=True)


def zipimages(zippathname, imagepath):
    zipdoc = zipfile.ZipFile(zippathname, 'w', zipfile.ZIP_DEFLATED)
    for d in os.listdir(imagepath):
        zipdoc.write(os.path.join(imagepath, d), os.path.split(d)[1], zipfile.ZIP_DEFLATED)
    zipdoc.close()


def getmypath():
    md5 = hashlib.md5()
    md5.update(str(uuid1()).encode('utf-8'))
    return os.path.join(os.getcwd(), "uploadfiles", md5.hexdigest())


def corpmargin(img):
    img = img[15:]
    img2 = img.sum(axis=2)
    imgleftfw = img[:, 1:]
    imgforleftfw = img[:, :-1]
    imgtopfw = img[1:]
    imgfortopfw = img[:-1]
    (row, col) = img2.shape
    rowdiffsum = abs(imgleftfw - imgforleftfw).sum(axis=1)
    coldiffsum = abs(imgtopfw - imgfortopfw).sum(axis=0)

    rowssum = img2.sum(axis=1)
    colssum = img2.sum(axis=0)
    row_top = 0
    raw_down = 0
    col_top = 0
    col_down = 0

    for r in range(0, row):
        if sum(rowdiffsum[r]) > 3 * 70 * col:
            # if 740 * col > rowssum[r] > 400 * col:
            row_top = r
            break

    for r in range(row - 1, 0, -1):
        if sum(rowdiffsum[r]) > 3 * 70 * col:
            # if 740 * col > rowssum[r] > 400 * col:
            raw_down = r
            break

    for c in range(0, col):
        if sum(coldiffsum[c]) > 3 * 70 * row:
            # if 740 * row > colssum[c] > 400 * row:
            col_top = c
            break

    for c in range(col - 1, 0, -1):
        if sum(coldiffsum[c]) > 3 * 70 * row:
            col_down = c
            break

    new_img = img[max(row_top - 30, 0):min(raw_down + 30, row), max(col_top - 30, 0):min(col_down + 30, col), 0:3]
    return new_img


if __name__ == '__main__':
    app.run()
