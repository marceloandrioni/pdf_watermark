# pdf_watermark
Add a watermark to a pdf file.

## CLI

In the command line just run: `pdf_watermark <input.pdf> <output.pdf> -wt|-wd|-wp <arg>`

e.g.:

* Add a watermark using simple text:

```
pdf_watermark.py examples/document.pdf examples/document_text.pdf -wt "Watermark from text"
```

* Add a watermark using a `docx` file as template:

```
pdf_watermark.py examples/document.pdf examples/document_docx.pdf -wd examples/watermark_docx.docx
```

* Add a watermark using a `pdf` file as template:

```
pdf_watermark.py examples/document.pdf examples/document_pdf.pdf  -wp examples/watermark_pdf.pdf
```

## GUI

Run the script with no arguments to open the GUI and then select the input and output pdf files.

<img src="./gui_example.png" alt="GUI" width="600"/>
