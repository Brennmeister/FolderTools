import qrcode
import json
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import PyPDF2
import tempfile

def create_img(data):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=16,
        border=4)

    qr.add_data(json.dumps(data))
    qr.make(fit=True)
    qrimg = qr.make_image(fill_color="black", back_color="white")

    # Generate Image
    label_size = (int(356/169*qrimg.size[0]), int(1*qrimg.size[0]))
    img = Image.new('RGB', label_size, (255, 255, 255))
    # Add QR-Code
    offset = (label_size[0]-qrimg.size[0], 0)
    img.paste(qrimg, offset)
    # Add Text
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(r"C:\Windows\Fonts\consola.ttf", 100)
    draw.text((30, int((label_size[1]-100)/2)), "sid{sid}".format(**data), (0, 0, 0), font=font)

    return img


class MyPDFWriter:
    """
    Platziert Etiketten auf einem Druckbogen
    """
    def __init__(self, ** kwargs):
        self.page_width_mm  = kwargs.get('page_width_mm', 210)
        self.page_height_mm = kwargs.get('page_height_mm', 297)
        self.out_file       = kwargs.get('out_file', None)

        self.pdfReader = None
        self.pdfWriter = PyPDF2.PdfFileWriter()

        self.num_labels_placed = 0
        self.place_positions_mm = None

        self.outPage = None

    def mm2units(self, length_mm):
        return length_mm * 2.54 / 161.3 * 180

    def units2mm(self, length_units):
        return length_units / self.mm2units(1)

    def add_empty_page(self):
        self.outPage = self.pdfWriter.addBlankPage(width=self.mm2units(self.page_width_mm),
                                              height=self.mm2units(self.page_height_mm))
        self.num_labels_placed = 0

    def write_pdf(self):
        fout = open(self.out_file,'wb')
        self.pdfWriter.write(fout)
        fout.close()
        return self.out_file

    def get_place_positions_mm(self):
        """
        Returns a list positions for the placement of the input pdf
        :param in_pdf: File Path of the PDF for which the placement positions should be calculated
        :return: a list of the placement positions
        """
        width = 35.6  # Breite/mm
        height = 16.9  # Höhe/mm
        gap_h = 2.5  # Lücke horizontal
        gap_v = 0.1  # Lücke vertikal

        offset = ((210 - 5 * width - 4 * gap_h) / 2, (297 - 14 * height - 13 * gap_v) / 2,)
        pos = list()
        for ii in range(0, 16):
            for jj in range(0, 5):
                pos.append(((offset[0] + (width + gap_h) * (jj)), (offset[1] + (height + gap_v) * (14-ii))))

        return pos

    def add_image(self, in_pdf):
        if self.pdfWriter.getNumPages()==0:
            self.add_empty_page()
        # Add label to current page. If page is full it adds another one
        if self.place_positions_mm is None:
            self.place_positions_mm = self.get_place_positions_mm()
        # Check if the page already is full and add an empty one
        if self.num_labels_placed>=len(self.place_positions_mm):
            # Add new empty Page
            self.add_empty_page()

        # Page is prepared for placing another label
        fin = open(in_pdf, 'rb')  # Open Input-File
        self.pdfReader = PyPDF2.PdfFileReader(fin)
        inPage = self.pdfReader.getPage(0)
        scale_factor = min(35.6 / self.units2mm(float(inPage.trimBox[2])), 16.9 / self.units2mm(float(inPage.trimBox[3])))

        pp_mm = self.place_positions_mm[self.num_labels_placed]
        self.outPage.mergeScaledTranslatedPage(inPage,
                                    scale=scale_factor,
                                    tx=0.0 + self.mm2units(pp_mm[0]),
                                    ty=0.0 + self.mm2units(pp_mm[1]))

        self.num_labels_placed += 1
        fin.close()




mpdf = MyPDFWriter()
im_list = list()
im_files = list()

tmp_file = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False).name
# mpdf.add_image(r'C:\Users\mpeter\Desktop\avery.pdf')
# mpdf.add_image(r'C:\Users\mpeter\Desktop\b.pdf')
first_run=True
offset_num = 18187
for ii in range(0+offset_num, 80+offset_num):
    data = {'sid': ii}
    img = create_img(data)

    img.save(tmp_file, 'PDF', resolution=200)

    im_list.append(img)
    im_files.append(tmp_file)
    mpdf.add_image(tmp_file)
    if first_run:
        mpdf.num_labels_placed=0
        mpdf.add_image(tmp_file)
        first_run=False


mpdf.out_file=r'C:\Users\mpeter\Desktop\test.pdf'
mpdf.write_pdf()