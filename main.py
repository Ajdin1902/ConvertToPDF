import os
import comtypes.client


wdFormatPDF = 17

inpath = "C:\\Users\\ajdin\\OneDrive\\MDZ\\All Rechnungen\\2025 Rechnungen\\August25\\"
outpath = "C:\\Users\\ajdin\\OneDrive\\MDZ\\All Rechnungen\\2025 Rechnungen\\August25\\PDF"



def convert_pdf(inpath_func, outpath_func):
    pdf_files = os.listdir(outpath_func)
    pdf_files = [pdf.replace("pdf", "docx") for pdf in pdf_files]
    [print(file) for file in pdf_files]
    invoiceFiles = os.listdir(inpath_func)
    for file in invoiceFiles:
        if file not in pdf_files:
            infile = inpath_func + "\\" + file  
            if infile.split(".")[-1] == "docx":
                print(file)
                print(infile)
                word = comtypes.client.CreateObject('Word.Application')
                doc = word.Documents.Open(infile)
                out_file_name = file.replace("docx", "pdf")
                out_file = outpath_func + "\\" + out_file_name
                doc.SaveAs(out_file, FileFormat=wdFormatPDF)
                print(f"{file} converted.\n")
                doc.Close()
                word.Quit()


if __name__ == '__main__':
    convert_pdf(inpath, outpath)
    files = os.listdir(path=inpath)
    print("The input folder:")
    try:
        [print(file.split("_")[1]) for file in files]
    except IndexError:
        print("Filename cannot be split.")
