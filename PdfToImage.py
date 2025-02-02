import os
import comtypes.client
from pdf2image import convert_from_path
from tkinter import filedialog, Tk

# Update this to your Poppler path
POPPLER_PATH = r"C:\Users\amari\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"


def select_file():
    """Opens a dialog to select a PDF or Word file."""
    root = Tk()
    root.withdraw()
    file_selected = filedialog.askopenfilename(
        title="Select a PDF or Word document",
        filetypes=[("PDF & Word Files", "*.pdf;*.docx;*.doc"),
                   ("PDF Files", "*.pdf"), ("Word Files", "*.docx;*.doc")]
    )
    return file_selected


def select_output_folder():
    """Opens a dialog to select an output folder."""
    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(
        title="Select a folder to save images")
    return folder_selected


def convert_word_to_pdf(word_path):
    """
    Converts a Word file to PDF.
    :param word_path: Path to the input Word file
    :return: Path to the converted PDF
    """
    if not os.path.exists(word_path):
        print(f"🚫 File not found: {word_path}")
        return None

    pdf_path = word_path.rsplit(".", 1)[0] + ".pdf"

    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False

    try:
        print(f"📄 Converting Word to PDF: {word_path}")
        doc = word.Documents.Open(os.path.abspath(word_path), ReadOnly=True)
        doc.SaveAs(pdf_path, FileFormat=17)  # FileFormat=17 is PDF
        doc.Close()
        print(f"✅ Converted {word_path} → {pdf_path}")
    except Exception as e:
        print(f"❌ Error converting Word to PDF: {e}")
        pdf_path = None
    finally:
        word.Quit()

    return pdf_path


def convert_pdf_to_images(pdf_path, output_folder, image_format="png", dpi=300):
    """
    Converts a PDF to images.
    :param pdf_path: Path to the input PDF file
    :param output_folder: Folder where images will be saved
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        print(f"🖼️ Converting PDF to images: {pdf_path}")
        images = convert_from_path(
            pdf_path, dpi=dpi, poppler_path=POPPLER_PATH)

        for i, image in enumerate(images, 1):
            image_path = os.path.join(
                output_folder, f"page_{i}.{image_format}")
            image.save(image_path, image_format.upper())
            print(f"✅ Saved: {image_path}")
    except Exception as e:
        print(f"❌ Failed to convert PDF to images: {e}")


def convert_word_to_images(word_path, output_folder):
    """
    Converts a Word document to images by first converting to PDF.
    :param word_path: Path to the Word file
    :param output_folder: Folder where images will be saved
    """
    pdf_path = convert_word_to_pdf(word_path)
    if pdf_path:
        convert_pdf_to_images(pdf_path, output_folder)


# Main Execution
file_path = select_file()
output_folder = select_output_folder()

if file_path and output_folder:
    if file_path.lower().endswith(".pdf"):
        convert_pdf_to_images(file_path, output_folder)
    elif file_path.lower().endswith((".doc", ".docx")):
        convert_word_to_images(file_path, output_folder)
    print("\n🎉 Conversion completed successfully!")
else:
    print("🚫 Operation cancelled.")
