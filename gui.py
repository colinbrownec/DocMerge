import os
import re
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QPushButton, QFileDialog, QListWidget, QListWidgetItem
from merge import merge_documents

class Filename():
  def __init__(self, name):
    self.name = name

  def __lt__(a, b):
    a_int = re.match('^\d+', a.name)
    b_int = re.match('^\d+', b.name)
    if a_int is not None and b_int is not None:
      return int(a_int.group(0)) < int(b_int.group(0))
    return a.name < b.name

def sort_filenames_with_prefix(files):
  # need to convert to list of objects with a custom comparison operator
  filenames = sorted([Filename(file) for file in files])
  return [filename.name for filename in filenames]

class TemplateList(QListWidget):
  def __init__(self, templates):
    super().__init__()
    self.checkboxes = []

    for template in templates:
      item = QListWidgetItem()
      item.setText(template)
      item.setCheckState(Qt.CheckState.Unchecked)

      self.checkboxes.append(item)
      self.addItem(item)

    # enable toggling item selection without needing to click the checkbox
    # self.itemClicked.connect(self._toggleItem)

  # returns the names of the selected templates
  def selectedTemplates(self):
    return [item.text() for item in self.checkboxes if item.checkState() == Qt.CheckState.Checked]

  def _toggleItem(self, item):
    item.setCheckState(Qt.CheckState.Checked if item.checkState() == Qt.CheckState.Unchecked else Qt.CheckState.Unchecked)

# Subclass QMainWindow to customize your application's main window
class DocumentMergeWindow(QMainWindow):
  def __init__(self, template_dir, master_docx='master.docx', default_filename='merge.docx'):
    super().__init__()
    self.template_dir = template_dir
    self.master_docx = master_docx
    self.default_filename = default_filename
    self.setWindowTitle('Document Merge')

    # find all Word documents in the template directory
    templates = []
    with os.scandir(self.template_dir) as it:
      for entry in it:
        if entry.name.endswith('docx') and entry.name != self.master_docx:
          templates.append(entry.name)

    templates = sort_filenames_with_prefix(templates)

    self.template_list = TemplateList(templates)
    self.build_btn = QPushButton('Build')
    self.build_btn.clicked.connect(self.buildDocument)

    layout = QVBoxLayout()
    layout.addWidget(self.template_list)
    layout.addWidget(self.build_btn)

    widget = QWidget()
    widget.setLayout(layout)
    self.setCentralWidget(widget)

  def buildDocument(self):
    filename, _ = QFileDialog.getSaveFileName(self, 'Save Merged Document', self.default_filename, 'Word Document (*docx)')
    if filename:
      templates = [os.path.join(self.template_dir, name) for name in self.template_list.selectedTemplates()]
      master = os.path.join(self.template_dir, self.master_docx)
      merge_documents(master, templates).save(filename)
