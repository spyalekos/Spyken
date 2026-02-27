import flet as ft
import fitz  # PyMuPDF
import docx
import edge_tts
import asyncio
import os
import tempfile

VOICE_MALE = "el-GR-NestorasNeural"
VOICE_FEMALE = "el-GR-AthinaNeural"

def extract_paragraphs(filepath: str) -> list[str]:
    ext = filepath.lower().split('.')[-1]
    paragraphs = []
    
    if ext == 'docx':
        doc = docx.Document(filepath)
        for p in doc.paragraphs:
            text = p.text.strip()
            if text:
                paragraphs.append(text)
    elif ext == 'pdf':
        doc = fitz.open(filepath)
        for page in doc:
            blocks = page.get_text("blocks")
            for block in blocks:
                text = block[4].strip()
                # Remove newlines mixed within the paragraph block (PyMuPDF blocks might contain newlines)
                text = " ".join(text.splitlines()).strip()
                if text:
                    paragraphs.append(text)
    else:
        raise ValueError("Μη υποστηριζόμενη μορφή αρχείου")
        
    return paragraphs

async def convert_to_audio(paragraphs: list[str], output_path: str, progress_callback):
    # temp_dir for storing partial files
    temp_dir = tempfile.mkdtemp()
    temp_files = []
    
    total = len(paragraphs)
    
    for i, p in enumerate(paragraphs):
        voice = VOICE_MALE if i % 2 == 0 else VOICE_FEMALE
        temp_file = os.path.join(temp_dir, f"part_{i}.mp3")
        
        communicate = edge_tts.Communicate(p, voice)
        await communicate.save(temp_file)
        
        temp_files.append(temp_file)
        progress_callback(i + 1, total)
        
    # Combine parts into single mp3 file
    with open(output_path, "wb") as outfile:
        for tf in temp_files:
            with open(tf, "rb") as infile:
                outfile.write(infile.read())
                
    # Cleanup
    for tf in temp_files:
        try:
            os.remove(tf)
        except Exception:
            pass
    try:
        os.rmdir(temp_dir)
    except Exception:
        pass


def main(page: ft.Page):
    page.title = "Spyken - Έγγραφο σε Ομιλία (MP3)"
    page.window.width = 650
    page.window.height = 700
    page.theme_mode = ft.ThemeMode.DARK
    page.padding = 30
    page.window.center()
    
    # UI Elements
    selected_files_list = ft.ListView(expand=1, spacing=10, padding=20, auto_scroll=True)
    status_text = ft.Text("Κατάσταση: Σε αναμονή", size=16, color=ft.colors.GREY_400)
    progress_bar = ft.ProgressBar(width=400, color="amber", bgcolor="#263238", value=0)
    progress_bar.visible = False
    
    file_queue = []

    def log(msg, error=False):
        color = ft.colors.RED_400 if error else ft.colors.GREEN_400
        selected_files_list.controls.append(ft.Text(msg, color=color))
        page.update()

    def on_file_picked(e: ft.FilePickerResultEvent):
        if e.files:
            for f in e.files:
                if f.path not in file_queue:
                    file_queue.append(f.path)
                    selected_files_list.controls.append(ft.Text(f"Επιλέχθηκε: {f.name}", color=ft.colors.WHITE70))
            page.update()

    file_picker = ft.FilePicker(on_result=on_file_picked)
    page.overlay.append(file_picker)
    
    def pick_files_clicked(e):
        file_picker.pick_files(allow_multiple=True, allowed_extensions=["pdf", "docx"])

    def clear_files(e):
        file_queue.clear()
        selected_files_list.controls.clear()
        selected_files_list.controls.append(ft.Text("Η λίστα καθαρίστηκε", color=ft.colors.GREY_400))
        page.update()

    async def start_conversion(e):
        if not file_queue:
            status_text.value = "Παρακαλώ επιλέξτε αρχεία πρώτα!"
            status_text.color = ft.colors.RED_400
            page.update()
            return

        convert_btn.disabled = True
        pick_btn.disabled = True
        clear_btn.disabled = True
        progress_bar.visible = True
        progress_bar.value = 0
        page.update()

        for filepath in file_queue:
            status_text.value = f"Εξαγωγή κειμένου: {os.path.basename(filepath)}"
            status_text.color = ft.colors.BLUE_400
            page.update()
            
            try:
                # Let the event loop breathe to update the UI
                await asyncio.sleep(0.1)
                
                # 1. Extract paragraphs
                paragraphs = extract_paragraphs(filepath)
                if not paragraphs:
                    log(f"Δεν βρέθηκε κείμενο στο {os.path.basename(filepath)}", error=True)
                    continue
                
                # 2. Convert to audio
                output_path = os.path.splitext(filepath)[0] + ".mp3"
                
                def update_progress(current, total):
                    progress_bar.value = current / total
                    status_text.value = f"Δημιουργία ήχου: {current}/{total} παράγραφοι"
                    page.update()

                await convert_to_audio(paragraphs, output_path, update_progress)
                
                log(f"Ολοκληρώθηκε: {os.path.basename(output_path)}")
                
            except Exception as ex:
                log(f"Σφάλμα στο {os.path.basename(filepath)}: {str(ex)}", error=True)

        status_text.value = "Κατάσταση: Όλες οι μετατροπές ολοκληρώθηκαν!"
        status_text.color = ft.colors.GREEN_400
        progress_bar.visible = False
        convert_btn.disabled = False
        pick_btn.disabled = False
        clear_btn.disabled = False
        page.update()

    # Buttons
    pick_btn = ft.ElevatedButton(
        "Επιλογή Αρχείων (.docx, .pdf)", 
        icon=ft.icons.FOLDER_OPEN,
        on_click=pick_files_clicked,
        style=ft.ButtonStyle(bgcolor=ft.colors.INDIGO_700, color=ft.colors.WHITE)
    )
    
    clear_btn = ft.ElevatedButton(
        "Καθαρισμός Λίστας", 
        icon=ft.icons.CLEAR,
        on_click=clear_files,
        style=ft.ButtonStyle(bgcolor=ft.colors.RED_900, color=ft.colors.WHITE)
    )

    convert_btn = ft.ElevatedButton(
        "Έναρξη Μετατροπής", 
        icon=ft.icons.PLAY_ARROW,
        on_click=start_conversion,
        style=ft.ButtonStyle(bgcolor=ft.colors.GREEN_700, color=ft.colors.WHITE)
    )

    # Layout assembling
    header = ft.Row(
        [
            ft.Icon(ft.icons.AUDIOTRACK, size=40, color=ft.colors.AMBER_400),
            ft.Text("Spyken", size=32, weight=ft.FontWeight.BOLD, color=ft.colors.AMBER_400),
            ft.Text("Έγγραφο σε MP3", size=20, color=ft.colors.GREY_300)
        ],
        alignment=ft.MainAxisAlignment.CENTER,
    )

    button_row = ft.Row([pick_btn, clear_btn], alignment=ft.MainAxisAlignment.CENTER)
    
    main_container = ft.Container(
        content=ft.Column(
            [
                header,
                ft.Divider(height=20, color="transparent"),
                button_row,
                ft.Container(
                    content=selected_files_list,
                    border=ft.border.all(1, ft.colors.WHITE24),
                    border_radius=10,
                    padding=10,
                    expand=True
                ),
                ft.Divider(height=10, color="transparent"),
                ft.Row([convert_btn], alignment=ft.MainAxisAlignment.CENTER),
                ft.Divider(height=10, color="transparent"),
                ft.Column([status_text, progress_bar], horizontal_alignment=ft.CrossAxisAlignment.CENTER)
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        ),
        expand=True
    )

    page.add(main_container)

if __name__ == "__main__":
    ft.app(target=main)
