from PyPDF2 import PdfFileWriter, PdfFileReader
import io
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from datetime import datetime
from PIL import Image
packet = io.BytesIO()
file = r'C:\Users\RHun\PycharmProjects\Material-Characterization-Forms\Material Characterization Form Data.xlsx'

# Make image have transparent background

Signature = 'Signature.jpg'
Checkmark = 'checkmark.png'

def process_image(image):

    img = Image.open(image)
    img = img.convert("RGBA")

    pixdata = img.load()

    for y in range(img.size[1]):
        for x in range(img.size[0]):
            if pixdata[x, y] == (255, 255, 255, 255):
                pixdata[x, y] = (255, 255, 255, 0)

    img.save("checkmark.PNG", "PNG")

# Process information from tracker
def process_form(file):

    # Create Dictionary containing locations of fields on page
    page_locations = {'Generator Name': (100, 690), 'Address': (100, 676),
                  'CityProvPostal': (100, 664), 'NameWaste': (100, 622), 'Colour': (60, 553), 'Odour': (181, 553),
                  'pH': (262, 553), 'FlashPoint': (369, 553), 'Component1': (21, 256), 'Component2': (21, 243),
                  'Component3': (21, 230), 'CAS1': (202, 256), 'CAS2': (202, 243), 'CAS3': (202, 230),
                  'Composition1': (444, 256), 'Composition2': (444, 243), 'Composition3': (444, 230),
                  'Name': (62, 142), 'Title': (405, 142), 'Company': (81, 129)}

    # Import data from tracker and
    Form_df = pd.read_excel(file, sheetname='Form')
    column_names = list(Form_df.columns.values)
    my_dict = {}
    for i in column_names:
        my_dict[i] = Form_df[i][0]

    # Merge text to input with location of fields
    all_info = {}
    for i,v in my_dict.items():
        for j,w in page_locations.items():
            if i == j:
                all_info[i] = [v, w]

    return all_info

# create new PDF with Reportlab
def create_pdf(data):

    c = canvas.Canvas(packet, pagesize=letter)

    # input data into input fields on pdf
    for i in data.keys():
        c.setFont("Helvetica", 8.5)
        try:
            if data[i][0] != data[i][0]:
                c.drawString(data[i][1][0], data[i][1][1], "")
            else:
                c.drawString(data[i][1][0], data[i][1][1], str(data[i][0]))
        except Exception as e:
            pass

    c.drawString(425, 676, "Nancy Varga")
    c.drawString(425, 664, "Nancy.Varga@LifeLabs.com")
    c.drawString(100, 651, "604-565-0043")
    c.drawString(310, 651, "45845")
    c.drawString(30, 597, "Lab Testing Wastes")
    c.drawString(405, 159, str(datetime.today().strftime("%m/%d/%Y")))

    # liquid pourable
    c.drawImage('checkmark.png', 15, 510, width=50, height=50, mask="auto")

    c.drawImage('Signature.PNG', 81, 94, width=62, height=34, mask="auto")

    c.save()

    # move to beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)

    # read in existing PDF
    existing_pdf = PdfFileReader(open("Material Characterization Form Template.pdf", "rb"))
    output = PdfFileWriter()

    # merge text with existing PDF
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)

    # write "output" to new file
    outputStream = open("test.pdf", "wb")
    output.write(outputStream)
    outputStream.close()

create_pdf(process_form(file))
