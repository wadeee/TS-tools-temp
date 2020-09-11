import hashlib
import os
import zipfile
from uuid import uuid1

import cv2
from flask import Flask, render_template, request, send_file
from skimage import io

app = Flask(__name__)


@app.route('/')
def index():
    return render_template("index.html")


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
