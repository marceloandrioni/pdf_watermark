#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Add a watermark to a pdf file.

Author: Marcelo Andrioni
https://github.com/marceloandrioni

"""

import os
import sys
import shutil
import tempfile
import argparse
import getpass
from pathlib import Path
import datetime
import numpy as np
from docx import Document
from docx.shared import RGBColor, Mm, Pt
from gooey import Gooey, GooeyParser
from subprocess import Popen
from pikepdf import Pdf, Page

# the docx conversion to pdf can be done with docx2pdf/word (windows only) or
# using libreoffice (windows and linux)
HAS_LIBREOFFICE = True if shutil.which('libreoffice') else False
try:
    import docx2pdf
    HAS_DOCX2PDF = True
except:
    HAS_DOCX2PDF = False

if not any((HAS_LIBREOFFICE, HAS_DOCX2PDF)):
    raise ValueError('libreoffice or docx2pdf/word must be installed.')


# Use CLI (instead of GUI) if the CLI arguments were passed.
# https://github.com/chriskiehl/Gooey/issues/449#issuecomment-534056010
if len(sys.argv) > 1:
    if '--ignore-gooey' not in sys.argv:
        sys.argv.append('--ignore-gooey')


def convert_to_pdf(input_docx, out_folder):
    """Convert docx file to pdf using libreoffice.
    Reference: https://stackoverflow.com/a/56067358/9707202
    Author: https://stackoverflow.com/users/7037499/dfresh22
    """

    # get libreoffice executable location
    libre_office = shutil.which('libreoffice')
    if not libre_office:
        raise ValueError('Could not find libreoffice executable location in path.')

    p = Popen([libre_office, '--headless', '--convert-to', 'pdf', '--outdir',
               out_folder, input_docx])
    # print([libre_office, '--convert-to', 'pdf', input_docx])
    p.communicate()


def create_watermark(wmark_pdf, text):

    wmark_pdf = Path(wmark_pdf)
    wmark_docx = wmark_pdf.with_suffix('.docx')

    # create docx
    document = Document()

    # set size as A4
    # https://stackoverflow.com/a/54757281/9707202
    for section in document.sections:
        section.page_height = Mm(297)
        section.page_width = Mm(210)
        section.left_margin = Mm(25.4)
        section.right_margin = Mm(25.4)
        section.top_margin = Mm(25.4)
        section.bottom_margin = Mm(25.4)
        section.header_distance = Mm(12.7)

        # set a very low footer
        section.footer_distance = Mm(5)   # Mm(12.7)

    # p = document.add_paragraph()
    p = document.sections[0].footer.add_paragraph()   # add the text as footer

    run = p.add_run(text)
    font = run.font
    font.name = 'Arial'
    font.size = Pt(12)
    font.color.rgb = RGBColor(255, 0, 0)

    document.save(wmark_docx)

    # convert docx to pdf
    if HAS_LIBREOFFICE:
        convert_to_pdf(wmark_docx, wmark_pdf.parent)
    else:
        docx2pdf.convert(wmark_docx, wmark_pdf)

    # remove docx
    os.remove(wmark_docx)

    return


def user_args():

    epilog = ("Example: "
              f"{sys.argv[0]} "
              "infile.pdf "
              "outfile.pdf "
              "-w 'This is my nice watermark !!!'")

    description = 'Add a watermark to a pdf file.'
    parser = GooeyParser(description=description,
                         epilog=epilog,
                         allow_abbrev=False)

    parser.add_argument('infile',
                        type=lambda x: Path(x),
                        help='Input pdf file.',
                        widget='FileChooser',
                        gooey_options={
                            'wildcard': 'PDF file (*.pdf)|*.pdf',
                            'message': 'Select input pdf file'})

    parser.add_argument('outfile',
                        type=lambda x: Path(x),
                        help='Output pdf file.',
                        widget='FileSaver',
                        gooey_options={
                            'wildcard': 'PDF file (*.pdf)|*.pdf',
                            'message': 'Select output pdf file'})

    parser.add_argument('-w', '--watermark',
                        required=True,
                        help='{:%Y-%m-%d %H:%M:%S}: {}'.format(
                            datetime.datetime.now(),
                            getpass.getuser()))

    args = parser.parse_args()

    if [args.infile.suffix, args.outfile.suffix] != ['.pdf', '.pdf']:
        raise argparse.ArgumentTypeError('Input/Output files must be pdf files.')

    if not args.infile.exists():
        raise argparse.ArgumentTypeError(f"Input file '{args.infile}' does "
                                         "not exist.")

    if args.outfile.exists() and os.path.samefile(args.infile, args.outfile):
        raise argparse.ArgumentTypeError("Input/Output files can't be the same.")

    return args


@Gooey(required_cols=1,
       progress_regex=r"^Page (?P<current>\d+)/(?P<total>\d+)$",
       progress_expr="current / total * 100")
def main():

    args = user_args()

    # create water mark file in temporary folder to overlay in main pdf file
    wmark_pdf = (Path(tempfile.gettempdir())
                 / 'watermark_{:05d}.pdf'.format(np.random.randint(10_000)))
    create_watermark(wmark_pdf, args.watermark)

    wmark = Pdf.open(wmark_pdf)
    thumbnail = Page(wmark.pages[0])

    print(f'Input file: {args.infile}')

    pdf = Pdf.open(args.infile)

    # merge water mark file with each page of input pdf file
    for page in pdf.pages:

        print(f'Page {page.index + 1}/{len(pdf.pages)}')

        page.add_overlay(thumbnail)

    print(f'Output file: {args.outfile}')
    pdf.save(args.outfile)

    pdf.close()

    # remove water mark file
    wmark.close()
    os.remove(wmark_pdf)

    print('Done!')


if __name__ == '__main__':
    main()
