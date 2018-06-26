import os
import zipfile
import xml.etree.cElementTree as ET

def main():
    countfiles()
    #   extractZip()
    #   parseXML()


def countfiles():
    docx_counter = 0
    pptx_counter = 0
    for root, dirs, files in os.walk('Test folder'):
        for file in files:
            print(file)
            if str.lower(file).endswith('.docx'):
                docx_counter += 1
            if str.lower(file).endswith('.pptx'):
                pptx_counter += 1

    print('.docx counter is ' + str(docx_counter))
    print('.pptx counter is ' + str(pptx_counter))
    return


def extractZip():
    my_zip = zipfile.ZipFile('Test Folder/Fonts.docx', 'r')
    with my_zip as zip_file:
        for member in zip_file.namelist():
            filename = os.path.basename(member)
            if filename == 'document.xml':
                source = zip_file.open(member)
                zip_file.extract(member)


    my_zip.close()
    return

def parseXML():
    tree = ET.parse('word/document.xml')
    root = tree.getroot()
    for child in root:
        print(child.tag, child.attrib)
    return

if __name__ == "__main__":
    main()
