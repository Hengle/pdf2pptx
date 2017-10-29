from pptx import Presentation
from pptx.util import Inches
import sys
import os
import shutil


def main():
    fpath = sys.argv[1]
    out_file = sys.argv[2]
    tmp_dir = 'tmp'
    while(True):
        if os.path.exists('./' + tmp_dir):
            tmp_dir = tmp_dir + '_'
        else:
            break
    os.mkdir('./' + tmp_dir)
    tmp_img = './' + tmp_dir + '/tmp.png'
    os.system('convert ' + fpath + ' ' + tmp_img)
    pics_len = len(os.listdir('./' + tmp_dir))
    pics = ['./tmp-' + str(n) + '.png' for n in range(pics_len)]
    prs = Presentation()
    for pic in pics:
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        left = top = Inches(0)
        left = Inches(0)
        pic = slide.shapes.add_picture('./' + tmp_dir + '/' + pic, left, top)
    prs.save(out_file)
    shutil.rmtree('./' + tmp_dir)


if __name__ == "__main__":
    main()
