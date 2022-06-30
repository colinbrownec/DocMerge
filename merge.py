from docx import Document
from docxcompose.composer import Composer

# see docs for docxcompose
# https://github.com/4teamwork/docxcompose

def merge_documents(root, documents):
  master = Document(root)
  composer = Composer(master)

  for doc in documents:
    composer.append(Document(doc))

  return composer
