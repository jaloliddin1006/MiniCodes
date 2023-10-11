import datetime
from aspose.pdf import ExcelSaveOptions, Document
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
import shutil
import os

class ConvertPdfToExcel:
    def __init__(self, base):
        self.BasePath = base
        self.ALLEXCELDATA = self.BasePath+"ALL_EXCEL_DATA.xlsx"
        self.NewPDFfiles = self.BasePath+"NEW_PDFS\\"
        
        files = os.listdir(base)

        new_pds = os.listdir(self.NewPDFfiles)

        for file_name in new_pds:
            print(file_name)
    
            self.startProgram(file_name)
            # break

    def startProgram(self, file_name):
        self.infile = self.NewPDFfiles + file_name
        self.client_code =self.infile.split(" ")[0]
        self.file_name = file_name


        if self.client_code == "10127":
            self.client_name = "FORVIA_10127"
        elif self.client_code == "10110":
            self.client_name = "KATCON_10110"

        self.PDFDIR = self.BasePath+f"{self.client_name}\\"
        self.outfile =  self.PDFDIR + self.client_name+"_all_data.xlsx"

        
        self.all_pdfs = []
        self.all_excels = []

        print(self.PDFDIR)
        print(self.outfile)

        self.PdfCutEach4Page()

    def create_folders(self):
        try:
            os.mkdir("pdfs_cut")
            os.mkdir("excels_cut")
            print("papkalar yaratildi")
        except Exception as err:
            # print(err)
            self.delete_folders()
            self.create_folders()
            pass

    def delete_folders(self):
        try:
            shutil.rmtree("pdfs_cut")
            shutil.rmtree("excels_cut")
            print("papkalar o'chirildi")
        except Exception as err:
            # print(err)
            pass

    def PdfMoveDirectory(self):
        file_to_copy = self.infile

        destination_directory = self.PDFDIR

        shutil.copy(file_to_copy, destination_directory)
        print("Fayl ko'chirildi")

    def NewPdfDelete(self):
        os.remove(self.infile)
        print("Fayl o'chirildi")

    def PdfCutEach4Page(self):
        self.PdfMoveDirectory()
        self.create_folders()
        with open(self.infile, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            total_pages = len(pdf_reader.pages)
            
            for start_page in range(0, total_pages, 4):
                end_page = min(start_page + 3, total_pages - 1)
                
                pdf_writer = PdfWriter()
                one_pdf = f'pdfs_cut/part_{start_page + 1}_to_{end_page + 1}.pdf'

                for page_num in range(start_page, end_page + 1):
                    page = pdf_reader.pages[page_num]
                    pdf_writer.add_page(page)
                
                with open(one_pdf, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)
                self.all_pdfs.append(one_pdf)

        print("Tugallandi. Yangi PDF fayllar saqlandi.")
        self.convert_PDF_to_XLSX()

    def convert_PDF_to_XLSX(self ):
        for file in self.all_pdfs:
            path_infile =  file
            path_outfile =  file.replace("pdfs_cut/", 'excels_cut/').split(".")[0] + ".xlsx"
            self.all_excels.append(path_outfile)
            
            document = Document(path_infile)
            save_option = ExcelSaveOptions()
            save_option.uniform_worksheets = True
            # save_option.insert_blank_column_at_first = True
            document.save(path_outfile, save_option)

        print("excelga o'tkazish tugallandi. ")
        self.MergeExcelsPartsOnePart()

    def MergeExcelsPartsOnePart(self):
        all_data = pd.DataFrame()
        sheets = ["Sheet1","Sheet2","Sheet3","Sheet4"]
        file_names = self.all_excels

        for file in file_names:
            # one_excel = pd.DataFrame()
            for sheet in sheets:
                try:
                    df2 = pd.read_excel(file, sheet_name=sheet)
                    DocumentDate = df2.iloc[11, 9]
                    OrderNo = df2.iloc[12, 9]
                    # print(DocumentDate)
                    # print(OrderNo)
                    endrow = df2[df2[df2.columns[0]]=='Release codes explanation:'].index[0]
                    data_to_append = df2.iloc[16 : endrow, :]

                    data_to_append.insert(0, "ORDER_CODE", OrderNo)
                    data_to_append.insert(1, "ORDER_DATE", DocumentDate)

                    all_data = pd.concat([all_data, data_to_append])
                except Exception as err:
                    print(err)

        all_data.columns = [
            "ORDER_CODE",
            "ORDER_DATE",
            "No.",
            "Description",
            "Outstanding Quantity",
            "Cumulative Qty.",
            "Unit of Measure",
            "Shipment Date",
            "Planned Receipt Date",
            "Release code",
            "Document",
            "Quantity",
            "Date",
            "Komentarz",
        ]
        # print(all_data)


        ## ustunni o'chirish
        all_data.drop(columns=["Komentarz", "Date", "Description","Cumulative Qty.", 
                               "Unit of Measure", "Planned Receipt Date", "Release code", 
                               "Document", "Quantity"], inplace=True,axis=1)

        ## ustunlarni tartiblash
        # s =  all_data[["DocumentDate"]]
        # all_data[["DocumentDate"]] = all_data[["OrderNo"]]
        # all_data[["OrderNo"]] = s

        #ustun nomini almashtirish
        all_data.rename(columns={"No.": "PRODUCT_CODE", 
                                 "Shipment Date": "DELIVERY_DATE", 
                                 "Outstanding Quantity":"QUANTITY"}, inplace=True)
        
        today = datetime.datetime.today().strftime('%d.%m.%Y')
        all_data.insert(0, "IMPORT_DATE", today)
        all_data.insert(1, "CUSTOMER_CODE", 10110)
        all_data.insert(2, "FILE_NAME_PDF", self.file_name)

        all_new_data = all_data
        if os.path.lexists(self.outfile):
            old_data = pd.read_excel(self.outfile)
            all_new_data = pd.concat([old_data, all_data])

        with pd.ExcelWriter(self.outfile) as writer:
            all_new_data.to_excel(writer, sheet_name='Sheet1', index=False)
        print("Ma'lumotlar muvaffaqiyatli yozildi.")

        all_client_data = all_data
        if os.path.lexists(self.ALLEXCELDATA):
            old_data = pd.read_excel(self.outfile)
            all_client_data = pd.concat([old_data, all_data])

        with pd.ExcelWriter(self.ALLEXCELDATA) as writer:
            all_client_data.to_excel(writer, sheet_name='Sheet1', index=False)
        self.delete_folders()
        self.NewPdfDelete()


# users = os.path.lexists("C:\\Users\\ProWin\\Documents\\Network\\Commserver\\pcm_import\\NEW_PDFS\\10110 PCM week 42.pdf")
# print(users)
base = "C:\\Users\\ProWin\\Documents\\Network\\Commserver\\pcm_import\\"
s = ConvertPdfToExcel(base)
# # s.PdfMoveDirectory()
s.delete_folders()
s.PdfCutEach4Page()

# files = os.listdir(base)

# new_pds = os.listdir(base+"NEW_PDFS\\")

# for file in new_pds:
#     print(file)
#     s = ConvertPdfToExcel(base+"NEW_PDFS\\"+file)
#     s.PdfCutEach4Page()
#     # break


# print(os.listdir("../"))