import win32com.client
import os
from .file_service import clean_filename


def split_word_to_pdfs(word_path: str, output_folder: str, filenames: list) -> list:
    results = []

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(word_path)
        total_pages = doc.ComputeStatistics(2)

        os.makedirs(output_folder, exist_ok=True)

        chunk_index = 0
        for start_page in range(1, total_pages + 1, 2):
            end_page = min(start_page + 1, total_pages)

            if chunk_index < len(filenames) and filenames[chunk_index]:
                pdf_name = clean_filename(filenames[chunk_index])
            else:
                pdf_name = f"Resolucion_{start_page:04d}"

            pdf_path = os.path.join(output_folder, f"{pdf_name}.pdf")

            try:
                rng = doc.GoTo(What=1, Which=1, Count=start_page)
                start_pos = rng.Start

                if end_page < total_pages:
                    rng_end = doc.GoTo(What=1, Which=1, Count=end_page + 1)
                    end_pos = rng_end.Start
                else:
                    end_pos = doc.Content.End

                export_range = doc.Range(start_pos, end_pos)

                new_doc = word.Documents.Add()
                new_doc.Content.Text = ""

                export_range.Copy()
                new_doc.Content.Paste()

                new_doc.ExportAsFixedFormat(
                    pdf_path,
                    17,
                    False,
                    0,
                    0,
                    1,
                    end_page - start_page + 1,
                    0,
                    True,
                    True,
                    0,
                    True,
                    True,
                    False
                )

                new_doc.Close(False)

                if os.path.exists(pdf_path):
                    results.append({
                        "chunk": chunk_index,
                        "pages": f"{start_page}-{end_page}",
                        "filename": pdf_name,
                        "path": pdf_path,
                        "success": True,
                        "error": None
                    })
                else:
                    results.append({
                        "chunk": chunk_index,
                        "pages": f"{start_page}-{end_page}",
                        "filename": pdf_name,
                        "path": pdf_path,
                        "success": False,
                        "error": "PDF no generado"
                    })

            except Exception as e:
                results.append({
                    "chunk": chunk_index,
                    "pages": f"{start_page}-{end_page}",
                    "filename": pdf_name,
                    "path": pdf_path,
                    "success": False,
                    "error": str(e)
                })

            chunk_index += 1

        doc.Close(False)
        word.Quit()

    except Exception as e:
        try:
            word.Quit()
        except:
            pass
        results.append({
            "chunk": -1,
            "pages": "N/A",
            "filename": "N/A",
            "path": "",
            "success": False,
            "error": f"Error general: {str(e)}"
        })

    return results


def export_mht_to_pdf(mht_path: str, pdf_path: str) -> tuple:
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(mht_path, False, True)
        doc.ExportAsFixedFormat(pdf_path, 17)
        doc.Close(False)
        word.Quit()

        if os.path.exists(pdf_path):
            return True, None
        return False, "PDF no generado"

    except Exception as e:
        try:
            word.Quit()
        except:
            pass
        return False, str(e)
