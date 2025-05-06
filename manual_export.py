import os
import datetime
import traceback
import win32com.client
from helper import send_email, log_to_file


def convert_word_to_pdf(word_filename: str, pdf_filename: str):
    word_path = os.path.abspath(word_filename)
    output_dir = os.path.abspath("exports")
    os.makedirs(output_dir, exist_ok=True)
    pdf_path = os.path.join(output_dir, pdf_filename)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(word_path)
        doc.Fields.Update()
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()

        success_msg = f"The PDF Document was successfully exported to: {pdf_path}"
        send_email("PDF Export was Successful!", success_msg)
        log_to_file(success_msg)
    except Exception:
        error_msg = traceback.format_exc()
        send_email("Scheduled Export has failed.", f"An Error occured:\n\n{error_msg}")
        log_to_file(f"Error:\n{error_msg}")
    finally:
        word.Quit()


if __name__ == "__main__":
    print("Word to PDF Export Automation!")

    word_filename_input = input("Enter Word filename (e.g., template.docx): ").strip()

    while word_filename_input == "":
        print("Oh you did not provide anything, please try again.")
        word_filename_input = input(
            "Enter Word filename (e.g., template.docx): "
        ).strip()
    if not word_filename_input.lower().endswith(".docx"):
        word_filename_input += ".docx"

    default_pdf_name = f"Document_{datetime.datetime.today().strftime("%Y-%m-%d")}.pdf"
    pdf_output = input(
        f"Please enter the PDF output name (default: {default_pdf_name}): "
    ).strip()

    if not pdf_output:
        pdf_output = default_pdf_name
    elif not pdf_output.lower().endswith(".pdf"):
        pdf_output += ".pdf"

    convert_word_to_pdf(word_filename_input, pdf_output)
