import os
from pdf2image import convert_from_path
from tkinter import filedialog, Tk


def select_pdf_file():
    """Opens a dialog to select a PDF file."""
    root = Tk()
    root.withdraw()  # Hide the root window
    file_selected = filedialog.askopenfilename(
        title="Select a PDF file", filetypes=[("PDF Files", "*.pdf")])
    return file_selected


def select_output_folder():
    """Opens a dialog to select an output folder."""
    root = Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(
        title="Select a folder to save images")
    return folder_selected


def convert_pdf_to_images(pdf_path, output_folder, image_format="png", dpi=300):
    """
    Converts a PDF into images.

    :param pdf_path: Path to the input PDF file
    :param output_folder: Folder where the images will be saved
    :param image_format: Image format (default: PNG)
    :param dpi: Resolution (default: 300 DPI)
    """
    if not os.path.exists(pdf_path):
        print("❌ Error: PDF file not found.")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    print("\n📄 Converting PDF pages to images...")

    try:
        images = convert_from_path(pdf_path, dpi=dpi)
        for i, image in enumerate(images, 1):
            image_path = os.path.join(
                output_folder, f"page_{i}.{image_format}")
            image.save(image_path, image_format.upper())
            print(f"✅ Saved: {image_path}")

        print("\n🎉 Conversion completed successfully!")

    except Exception as e:
        print(f"❌ Failed to convert PDF to images: {e}")


# Select PDF and output folder
pdf_file = select_pdf_file()
output_folder = select_output_folder()

if pdf_file and output_folder:
    convert_pdf_to_images(pdf_file, output_folder)
else:
    print("🚫 Operation cancelled.")
