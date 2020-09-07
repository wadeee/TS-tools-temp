import os
import shutil
import zipfile

import pythoncom
from fitz import fitz
from flask import Flask, render_template, request, send_file
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

    docname = request.files.get('originFileA').filename
    docpathname = os.path.join(os.getcwd(), docname)
    namepre = os.path.splitext(docname)[0]
    pdfname = namepre + ".pdf"
    pdfpathname = os.path.join(os.getcwd(), pdfname)
    imagepath = os.path.join(os.getcwd(), namepre)
    zipfilename = namepre + '.zip'
    zippath = os.path.join(os.getcwd(), 'zipfiles')
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

    zipdoc = zipfile.ZipFile(zippathname, 'w', zipfile.ZIP_DEFLATED)
    for d in os.listdir(imagepath):
        zipdoc.write(os.path.join(imagepath, d), os.path.split(d)[1], zipfile.ZIP_DEFLATED)
    zipdoc.close()
    shutil.rmtree(imagepath)
    os.remove(pdfpathname)

    return send_file(zippathname,
                     mimetype='application/x-zip-compressed;charset=utf-8',
                     attachment_filename=zipfilename,
                     as_attachment=True)


@app.route('/b', methods=['POST'])
def b():
    pdfname = request.files.get('originFileB').filename
    pdfpathname = os.path.join(os.getcwd(), pdfname)
    namepre = os.path.splitext(pdfname)[0]

    imagepath = os.path.join(os.getcwd(), namepre)
    zipfilename = namepre + '.zip'
    zippath = os.path.join(os.getcwd(), 'zipfiles')
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

        pix.writePNG(os.path.join(imagepath, '{}.png'.format(pg)))
    filedoc.close()

    zipdoc = zipfile.ZipFile(zippathname, 'w', zipfile.ZIP_DEFLATED)
    for d in os.listdir(imagepath):
        zipdoc.write(os.path.join(imagepath, d), os.path.split(d)[1], zipfile.ZIP_DEFLATED)
    zipdoc.close()
    shutil.rmtree(imagepath)
    os.remove(pdfpathname)

    return send_file(zippathname,
                     mimetype='application/x-zip-compressed;charset=utf-8',
                     attachment_filename=zipfilename,
                     as_attachment=True)


if __name__ == '__main__':
    app.run()
