import os
import zipfile
import xml.etree.cElementTree as ET
import mimetypes


font_list = []


def main():
    countfiles()
    print(font_list)
    outfile = open("UsedFonts.txt", "w")
    for font in font_list:
        outfile.write(font + '\n')
    outfile.close()

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
            #   print(str(member))
            filename = os.path.basename(member)
            if filename == 'fontTable.xml':
                source = zip_file.read(member)
    my_zip.close()
    return source


def parsexml(sourcexml):
    tree = ET.ElementTree(ET.fromstring(sourcexml))
    root = tree.getroot()
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    for font in root.findall("w:font", namespaces):
        font_list.append((font.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name")))
    return


if __name__ == "__main__":
    main()