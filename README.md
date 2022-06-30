# Usage
``` cmd
python ./main.py
```

Or run `DocMerge.cmd`

# Setup
Python is required.
``` cmd
python -m pip install ./requirements.txt
```

## Template Directory
Any Word document found in the template directory will be presented as an option.

Word documents are sorted by name.
If the document name starts with an integer prefix, the document will be sorted with respect to other integer prefixes.

For example, the following items will be sorted in the below order,
```
/templates/1 - registration.docx
/templates/2 - software license.docx
/templates/3 - technical support.docx
/templates/4 - client hosted software.docx
```

## Master Document
A master document can be used to provide base styling (and content) to all merged documents.

An empty master document is provided. This is an new, empty Word document, but it has been edited outside of Word to remove the empty paragraph inserted into new Word documents.
