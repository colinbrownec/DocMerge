import os
import argparse
from PyQt6.QtWidgets import QApplication
from gui import DocumentMergeWindow

template_dir = os.path.dirname(os.path.realpath(__file__))
template_dir = os.path.join(template_dir, 'templates')

parser = argparse.ArgumentParser(description='Merge multiple Word document templates')
parser.add_argument('--templates', help='Directory containing Word documents to be used as templates', default=template_dir)
parser.add_argument('--master', help='Name of the master document, must be in the templates directory', default='master.docx')
args = parser.parse_args()

app = QApplication([])

window = DocumentMergeWindow(args.templates, args.master)
window.show()

app.exec()
