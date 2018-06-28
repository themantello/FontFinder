import os
import zipfile
import xml.etree.cElementTree as ET
import mimetypes


font_list = []


def main():
    countfiles()
    print(font_list)


def countfiles():
    for root, dirs, files in os.walk('Test folder'):
        for file in files:
            fileType = mimetypes.guess_type(file)[0]
            if str(fileType).startswith('application/vnd.openxml'):
                parsexml((extractzip(os.path.join(root, file))))
    return


def extractzip(xfile):
    my_zip = zipfile.ZipFile(xfile, 'r')
    with my_zip as zip_file:
        for member in zip_file.namelist():
            filename = os.path.basename(member)
            if filename == 'document.xml':
                source = zip_file.read(member)
    my_zip.close()
    return source


def parsexml(sourcexml):
    tree = ET.ElementTree(ET.fromstring(sourcexml))
    root = tree.getroot()
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    for font in root.findall("w:body/w:p/w:pPr/w:rPr/w:rFonts", namespaces):
        font_list.append((font.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii")))
    return


if __name__ == "__main__":
    main()