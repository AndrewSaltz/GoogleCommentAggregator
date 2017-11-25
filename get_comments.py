import zipfile
import pandas as pd
import lxml.etree as etree
from io import BytesIO
import datetime
from tkinter.filedialog import askopenfilename

ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'} # The Namespace - from stackoverflow
df = pd.DataFrame(columns=['Title','Author', 'Date', 'Comment']) # Set up the spreadsheet
nope = ['_Marked as resolved_', '_Re-opened_'] # We don't want these published. Change this if you'd like
total = 0
total_comments = 0
file = askopenfilename() # The jawn that lets git you select a file
with zipfile.ZipFile(file) as parent_zip:
    file_list = parent_zip.namelist() #List all files in the zip
    for f in file_list:
        zfiledata = BytesIO(parent_zip.read(f)) # .docx is a zip - use this to read zip in a zip
        with zipfile.ZipFile(zfiledata) as child_zip:
            try:
                document = child_zip.read('word/comments.xml') # Read the comments
                tree = etree.XML(document)
                comments = tree.xpath('//w:comment', namespaces=ooXMLns) # Find all the comments in that doc
                for c in comments:
                    title = f # Title of the Doc
                    author = str(c.xpath('@w:author', namespaces=ooXMLns))
                    publish = str(c.xpath('@w:date', namespaces=ooXMLns))
                    comment = c.xpath('string(.)', namespaces=ooXMLns)
                    author = author.replace('[','').replace(']','').replace("'",'') # Formatting (here and below)
                    publish = publish.replace('[','').replace(']','').replace("'",'')
                    if comment not in nope: # Sort rubbish
                        new_row = [title,author,publish,comment]
                        a = len(df)
                        df.loc[a] = new_row
                        total_comments += 1 # So as to always write one row down
            except KeyError:
                print("No Comments Here")
        total +=1
d = str(datetime.date.today())
writer = pd.ExcelWriter("Comments %s .xlsx" % d)
df.to_excel(writer, sheet_name="Comments",)
workbook = writer.book
worksheet = writer.sheets['Comments']
worksheet.set_column(1, 3, 20) # Fiddle with this if you'd like, just sets the columns on Excel
workbook.close()
print ("Found {0} comments from {1} files".format(total_comments, total))