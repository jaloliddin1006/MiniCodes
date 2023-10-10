from aspose.pdf import ExcelSaveOptions, Document
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
import shutil
import os

class ConvertPdfToExcel:
    def __init__(self, infile):
        self.infile = infile
        self.outfile = infile.replace(".pdf", ".xlsx")
        self.all_pdfs = []
        self.all_excels = []

    def create_folders(self):
        try:
            os.mkdir("pdfs_cut")
            os.mkdir("excels_cut")
            # print("papkalar yaratildi")
        except Exception as err:
            # print(err)
            self.delete_folders()
            self.create_folders()
            pass

    def delete_folders(self):
        try:
            shutil.rmtree("pdfs_cut")
            shutil.rmtree("excels_cut")
            # print("papkalar o'chirildi")
        except Exception as err:
            # print(err)
            pass

    def PdfCutEach4Page(self):
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

        # print("Tugallandi. Yangi PDF fayllar saqlandi.")
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

        # print("excelga o'tkazish tugallandi. ")
        self.MergeExcelsPartsOnePart()

    def MergeExcelsPartsOnePart(self):
        all_data = pd.DataFrame()
        sheets = ["Sheet1","Sheet2","Sheet3","Sheet4"]
        file_names = self.all_excels

        for file in file_names:
            one_excel = pd.DataFrame()
            for sheet in sheets:
                try:
                    df2 = pd.read_excel(file, sheet_name=sheet)
                    DocumentDate = df2.iloc[11, 9]
                    OrderNo = df2.iloc[12, 9]
                    # print(DocumentDate)
                    # print(OrderNo)
                    endrow = df2[df2[df2.columns[0]]=='Release codes explanation:'].index[0]
                    data_to_append = df2.iloc[16 : endrow, :]

                    data_to_append.insert(0, "DocumentDate", DocumentDate)
                    data_to_append.insert(1, "OrderNo", OrderNo)

                    all_data = pd.concat([all_data, data_to_append])
                except Exception as err:
                    print(err)

        all_data.columns = [
            "DocumentDate",
            "OrderNo",
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
        with pd.ExcelWriter(self.outfile) as writer:
            all_data.to_excel(writer, sheet_name='Sheet1', index=False)
        # print("Ma'lumotlar muvaffaqiyatli yozildi.")

        self.delete_folders()


s = ConvertPdfToExcel("test.pdf")
s.delete_folders()
s.PdfCutEach4Page()
