import re
from enums import control_char_table
from hwp5_table import HwpFile
import os
import shutil


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def gethering(dirname, list):
    filenames = os.listdir(dirname)
    for filename in filenames:
        full_filename = os.path.join(dirname, filename)
        if os.path.isdir(full_filename):
            gethering(full_filename, list)
        else:
            ext = os.path.splitext(full_filename)[-1]
            if ext == '.hwp':
                list.append(full_filename)

def listing(dirname) :
    list = []
    gethering(dirname, list)
    return list

def cli(input):
    def get_text(self):
        regex = re.compile(rb'([\x00-\x1f])\x00')

        text = ''

        cursor_idx = 0
        search_idx = 0

        while cursor_idx < len(self.payload):
            if search_idx < cursor_idx:
                search_idx = cursor_idx

            searched = regex.search(self.payload, search_idx)
            if searched:
                pos = searched.start()

                if pos & 1:
                    search_idx = pos + 1
                elif pos > cursor_idx:
                    text += self.payload[cursor_idx:pos].decode('utf-16')
                    cursor_idx = pos
                else:
                    control_char = ord(searched.group(1))
                    control_char_size = control_char_table[control_char][1].size

                    if control_char == 0x0a:
                        text += '\n'

                    cursor_idx = pos + control_char_size * 2
            else:
                text += self.payload[search_idx:].decode('utf-16')
                break

        return text

    hwp = HwpFile(input)

    section_idx = 0
    while hwp.ole.exists('BodyText/Section%d' % section_idx):
        r=hwp.get_record_tree(section_idx)
        section_idx += 1

    filename = os.path.basename(input)
    dirname = os.path.dirname(input)

    if (filename.find('[단가표]') == -1 and filename.find('[사양]') == -1 and filename.find('[견적서]') == -1 and filename.find('[발주서]') == -1 and filename.find('[검사성적서]') == -1 and filename.find('[매입출]') == -1) :
        firstLine = get_text(r.children[0].children[0]).replace(" ", "")
        if (firstLine.find('단가표') != -1) :
            shutil.move(input, dirname+'/[단가표]'+filename)
            print(dirname+'/[단가표]'+filename)
        elif (firstLine.find('사양') != -1 or firstLine.find('제작') != -1) :
            shutil.move(input, dirname + '/[사양]' + filename)
            print(dirname+'/[사양]'+filename)
        elif (firstLine.find('발주서') != -1) :
            shutil.move(input, dirname + '/[발주서]' + filename)
            print(dirname+'/[발주서]'+filename)
        elif (firstLine.find('견적서') != -1 or firstLine.find('견적요청서') != -1 ) :
            shutil.move(input, dirname + '/[견적서]' + filename)
            print(dirname+'/[견적서]'+filename)
        elif (firstLine.find('검사성적서') != -1):
            shutil.move(input, dirname + '/[검사성적서]' + filename)
            print(dirname + '/[검사성적서]' + filename)
        elif (firstLine.find('현대모터스') != -1 and firstLine.find(')') != -1):
            shutil.move(input, dirname + '/[매입출]' + filename)
            print(dirname + '/[매입출]' + filename)


if __name__ == '__main__':
    filelist = listing(BASE_DIR)
    for file in filelist:
        cli(file)