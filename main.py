import flet as ft
import fitz  # PyMuPDF
import docx
import edge_tts
import asyncio
import os
import tempfile
import textwrap

VOICE_MALE = "el-GR-NestorasNeural"
VOICE_FEMALE = "el-GR-AthinaNeural"
VOICE_EN_MALE = "en-GB-RyanNeural"
VOICE_EN_FEMALE = "en-GB-SoniaNeural"

# Video frame resolution
VIDEO_W = 1280
VIDEO_H = 720
VIDEO_FPS = 10  # Reduced from 24 to speed up rendering significantly


def is_english(text: str) -> bool:
    """Returns True if the majority of letters in text are Latin (English)."""
    latin = sum(1 for c in text if 'A' <= c <= 'Z' or 'a' <= c <= 'z')
    greek = sum(1 for c in text if '\u0370' <= c <= '\u03ff' or '\u1f00' <= c <= '\u1fff')
    return latin > greek


def is_valid_text(text: str) -> bool:
    """Returns True only if text has enough real words for TTS."""
    if len(text) < 3:
        return False
    if not any(c.isalpha() for c in text):
        return False
    return True


def chunk_text(text: str, max_chars: int = 800) -> list[str]:
    """Split text into chunks of max_chars, breaking at sentence boundaries."""
    if len(text) <= max_chars:
        return [text]
    chunks = []
    while len(text) > max_chars:
        split_pos = max_chars
        for sep in ('. ', '! ', '? ', ', '):
            pos = text.rfind(sep, 0, max_chars)
            if pos != -1:
                split_pos = pos + len(sep)
                break
        chunks.append(text[:split_pos].strip())
        text = text[split_pos:].strip()
    if text:
        chunks.append(text)
    return chunks


def extract_paragraphs(filepath: str) -> list[str]:
    ext = filepath.lower().split('.')[-1]
    paragraphs = []

    if ext == 'docx':
        doc = docx.Document(filepath)
        for p in doc.paragraphs:
            text = p.text.strip()
            if is_valid_text(text):
                paragraphs.append(text)
    elif ext == 'pdf':
        doc = fitz.open(filepath)
        for page in doc:
            blocks = page.get_text("blocks")
            for block in blocks:
                text = " ".join(block[4].strip().splitlines()).strip()
                if is_valid_text(text):
                    paragraphs.append(text)
    else:
        raise ValueError("Μη υποστηριζόμενη μορφή αρχείου")

    return paragraphs


# ─────────────────────────── VIDEO HELPERS ────────────────────────────────────

def extract_paragraphs_pdf_with_pos(filepath: str) -> list[tuple]:
    """
    Returns list of (text, page_idx, fitz.Rect) for each valid paragraph.
    """
    results = []
    doc = fitz.open(filepath)
    for page_idx, page in enumerate(doc):
        blocks = page.get_text("blocks")
        for block in blocks:
            # block = (x0, y0, x1, y1, "text", block_no, block_type)
            text = " ".join(block[4].strip().splitlines()).strip()
            if is_valid_text(text):
                rect = fitz.Rect(block[0], block[1], block[2], block[3])
                results.append((text, page_idx, rect, doc))
    return results


def render_page_pdf_image(pdf_doc, page_idx: int, highlight_rect=None, target_w: int = VIDEO_W, target_h: int = VIDEO_H):
    """
    Render a PDF page as a PIL RGBA image sized to fit target_w x target_h.
    Optionally draw a highlight rectangle over `highlight_rect` (fitz.Rect).
    Returns a PIL Image (RGB).
    """
    from PIL import Image, ImageDraw

    page = pdf_doc[page_idx]
    page_rect = page.rect
    scale = min(target_w / page_rect.width, target_h / page_rect.height)
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, alpha=False)

    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    # Centre the page on a dark canvas
    canvas = Image.new("RGB", (target_w, target_h), (30, 30, 40))
    x_off = (target_w - img.width) // 2
    y_off = (target_h - img.height) // 2
    canvas.paste(img, (x_off, y_off))

    if highlight_rect is not None:
        overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        hx0 = int(highlight_rect.x0 * scale) + x_off
        hy0 = int(highlight_rect.y0 * scale) + y_off
        hx1 = int(highlight_rect.x1 * scale) + x_off
        hy1 = int(highlight_rect.y1 * scale) + y_off
        # Semi-transparent yellow fill
        draw.rectangle([hx0, hy0, hx1, hy1], fill=(255, 220, 0, 80))
        # Solid orange border
        draw.rectangle([hx0, hy0, hx1, hy1], outline=(255, 140, 0, 230), width=3)
        canvas = canvas.convert("RGBA")
        canvas = Image.alpha_composite(canvas, overlay).convert("RGB")

    return canvas


def render_docx_paragraph_image(text: str, para_idx: int, total: int, target_w: int = VIDEO_W, target_h: int = VIDEO_H):
    """
    Render a DOCX paragraph as a PIL Image (fallback renderer — no Word needed).
    Shows the paragraph text with a highlighted background box, and a header
    showing current / total paragraphs.
    Returns a PIL Image (RGB).
    """
    from PIL import Image, ImageDraw, ImageFont

    BG = (24, 28, 42)
    HEADER_BG = (40, 44, 64)
    HIGHLIGHT = (255, 220, 0)
    TEXT_COLOR = (30, 30, 30)
    HEADER_COLOR = (180, 180, 200)

    canvas = Image.new("RGB", (target_w, target_h), BG)
    draw = ImageDraw.Draw(canvas)

    # Header bar
    draw.rectangle([0, 0, target_w, 60], fill=HEADER_BG)
    header_txt = f"Παράγραφος {para_idx + 1} / {total}"
    try:
        header_font = ImageFont.truetype("arial.ttf", size=24)
        draw.text((target_w // 2, 30), header_txt, font=header_font, fill=HEADER_COLOR, anchor="mm")
    except Exception:
        draw.text((target_w // 2, 30), header_txt, fill=HEADER_COLOR, anchor="mm")

    # Figure out a reasonable font size that fits the text box
    box_w = target_w - 120
    box_h = target_h - 160
    box_x = 60
    box_y = 90

    # Word-wrap
    max_chars_per_line = max(30, box_w // 14)
    lines = textwrap.wrap(text, width=max_chars_per_line)
    if not lines:
        lines = ["(κενή παράγραφος)"]

    # Estimate line height and adjust font size to fit
    line_h = box_h // max(len(lines), 1)
    line_h = min(line_h, 50)
    line_h = max(line_h, 22)

    # Draw highlight box
    text_block_h = len(lines) * line_h + 40
    text_block_w = box_w + 40
    tx = box_x - 20
    ty = box_y - 10
    draw.rectangle([tx, ty, tx + text_block_w, ty + text_block_h], fill=HIGHLIGHT)
    draw.rectangle([tx, ty, tx + text_block_w, ty + text_block_h], outline=(255, 140, 0), width=4)

    # Draw text lines
    y_cursor = box_y + 10
    try:
        font_size = max(12, int(line_h * 0.7))
        main_font = ImageFont.truetype("arial.ttf", size=font_size)
    except Exception:
        main_font = ImageFont.load_default()

    for line in lines:
        draw.text((box_x + 10, y_cursor), line, font=main_font, fill=TEXT_COLOR)
        y_cursor += line_h
        if y_cursor > box_y + box_h:
            break

    # Footer
    draw.rectangle([0, target_h - 40, target_w, target_h], fill=HEADER_BG)
    footer_txt = "Spyken · MP4 by spyalekos"
    try:
        footer_font = ImageFont.truetype("arial.ttf", size=18)
        draw.text((target_w // 2, target_h - 20), footer_txt, font=footer_font, fill=(100, 100, 130), anchor="mm")
    except Exception:
        draw.text((target_w // 2, target_h - 20), footer_txt, fill=(100, 100, 130), anchor="mm")

    return canvas


def get_mp3_duration(path: str) -> float:
    """Return duration of an mp3 file in seconds."""
    try:
        from mutagen.mp3 import MP3
        audio = MP3(path)
        return audio.info.length
    except Exception:
        return 3.0  # fallback


async def generate_tts_chunk(text: str, voice: str, out_path: str) -> bool:
    """Generate a single TTS mp3 chunk. Returns True on success."""
    for attempt in range(3):
        try:
            communicate = edge_tts.Communicate(text, voice)
            await communicate.save(out_path)
            if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                return True
        except Exception:
            await asyncio.sleep(0.5)
    return False


async def convert_to_video(filepath: str, output_path: str, progress_callback):
    """
    Master function: extract paragraphs (with positions for PDF),
    generate TTS per paragraph, build video frames, assemble mp4.
    """
    import numpy as np
    from moviepy import ImageClip, AudioFileClip, concatenate_videoclips, concatenate_audioclips
    from proglog import ProgressBarLogger

    class FletMoviepyLogger(ProgressBarLogger):
        def __init__(self, cb):
            super().__init__()
            self.ui_callback = cb

        def bars_callback(self, bar, attr, value, old_value=None):
            # 't' is the main progress bar for write_videofile
            if bar == 't':
                total = self.bars[bar].get('total', 1)
                if total > 0:
                    pct = int((value / total) * 100)
                    self.ui_callback(value, total, f"Συναρμολόγηση βίντεο... {pct}%")

    ext = filepath.lower().split('.')[-1]
    temp_dir = tempfile.mkdtemp()

    try:
        # ── 1. Extract paragraphs ─────────────────────────────────────────────
        if ext == 'pdf':
            pdf_doc = fitz.open(filepath)
            para_data = []  # (text, page_idx, rect)
            for page_idx, page in enumerate(pdf_doc):
                blocks = page.get_text("blocks")
                for block in blocks:
                    text = " ".join(block[4].strip().splitlines()).strip()
                    if is_valid_text(text):
                        rect = fitz.Rect(block[0], block[1], block[2], block[3])
                        para_data.append((text, page_idx, rect))
        else:  # docx
            doc_obj = docx.Document(filepath)
            para_data = []
            for p in doc_obj.paragraphs:
                text = p.text.strip()
                if is_valid_text(text):
                    para_data.append((text, len(para_data), None))

        if not para_data:
            raise ValueError("Δεν βρέθηκαν παράγραφοι στο αρχείο.")

        total = len(para_data)
        clips = []
        voice_index = 0

        # ── 2. Per-paragraph: TTS + frame ─────────────────────────────────────
        for i, para_item in enumerate(para_data):
            text = para_item[0]
            progress_callback(i, total, f"Παράγραφος {i+1}/{total}: TTS…")
            await asyncio.sleep(0)  # yield to UI

            # Pick voice
            voice = (
                (VOICE_EN_MALE if voice_index % 2 == 0 else VOICE_EN_FEMALE)
                if is_english(text)
                else (VOICE_MALE if voice_index % 2 == 0 else VOICE_FEMALE)
            )

            # For long paragraphs, generate TTS for all chunks and concatenate their audio
            chunks = chunk_text(text, 800)
            chunk_audio_clips = []
            
            for c_idx, chunk in enumerate(chunks):
                chunk_audio_path = os.path.join(temp_dir, f"audio_{i}_{c_idx}.mp3")
                ok = await generate_tts_chunk(chunk, voice, chunk_audio_path)
                if ok:
                    chunk_audio_clips.append(AudioFileClip(chunk_audio_path))

            if chunk_audio_clips:
                if len(chunk_audio_clips) == 1:
                    combined_audio = chunk_audio_clips[0]
                else:
                    combined_audio = concatenate_audioclips(chunk_audio_clips)
                duration = combined_audio.duration
                voice_index += 1
            else:
                combined_audio = None
                duration = 3.0

            # ── 3. Render frame ───────────────────────────────────────────────
            progress_callback(i, total, f"Παράγραφος {i+1}/{total}: Εικόνα…")
            await asyncio.sleep(0)

            if ext == 'pdf':
                _, page_idx, rect = para_item
                frame_img = render_page_pdf_image(pdf_doc, page_idx, highlight_rect=rect)
            else:
                _, para_idx, _ = para_item
                frame_img = render_docx_paragraph_image(text, para_idx, total)

            frame_np = np.array(frame_img)

            # ── 4. Build clip ─────────────────────────────────────────────────
            img_clip = ImageClip(frame_np, duration=duration)

            if combined_audio:
                # Need to explicitly set duration on the audio clip, 
                # though concatenate_audioclips does this usually.
                if combined_audio.duration > duration:
                    combined_audio = combined_audio.subclipped(0, duration)
                img_clip = img_clip.with_audio(combined_audio)

            clips.append(img_clip)

        # ── 5. Assemble final video ────────────────────────────────────────────
        progress_callback(total, total, "Συναρμολόγηση βίντεο...")
        await asyncio.sleep(0)

        ui_logger = FletMoviepyLogger(progress_callback)

        final = concatenate_videoclips(clips, method="compose")
        final.write_videofile(
            output_path,
            fps=VIDEO_FPS,
            codec="libx264",
            audio_codec="aac",
            temp_audiofile=os.path.join(temp_dir, "tmp_audio.m4a"),
            remove_temp=True,
            logger=ui_logger,
        )
        final.close()

    finally:
        # Cleanup temp files
        for f in os.listdir(temp_dir):
            try:
                os.remove(os.path.join(temp_dir, f))
            except Exception:
                pass
        try:
            os.rmdir(temp_dir)
        except Exception:
            pass


# ──────────────────────────── AUDIO CONVERSION ────────────────────────────────

async def convert_to_audio(paragraphs: list[str], output_path: str, progress_callback):
    temp_dir = tempfile.mkdtemp()
    temp_files = []

    all_chunks = []
    for p in paragraphs:
        all_chunks.extend(chunk_text(p))

    total = len(all_chunks)
    voice_index = 0

    for i, chunk in enumerate(all_chunks):
        voice = (
            (VOICE_EN_MALE if voice_index % 2 == 0 else VOICE_EN_FEMALE)
            if is_english(chunk)
            else (VOICE_MALE if voice_index % 2 == 0 else VOICE_FEMALE)
        )
        temp_file = os.path.join(temp_dir, f"part_{i}.mp3")

        success = False
        for attempt in range(3):
            try:
                communicate = edge_tts.Communicate(chunk, voice)
                await communicate.save(temp_file)
                if os.path.exists(temp_file) and os.path.getsize(temp_file) > 0:
                    success = True
                    break
            except Exception:
                await asyncio.sleep(0.5)

        if success:
            temp_files.append(temp_file)
            voice_index += 1

        progress_callback(i + 1, total)

    with open(output_path, "wb") as outfile:
        for tf in temp_files:
            with open(tf, "rb") as infile:
                outfile.write(infile.read())

    for tf in temp_files:
        try:
            os.remove(tf)
        except Exception:
            pass
    try:
        os.rmdir(temp_dir)
    except Exception:
        pass


# ──────────────────────────────── UI ──────────────────────────────────────────

def main(page: ft.Page):
    APP_VERSION = "0.93"
    page.title = "Spyken by spyalekos - Έγγραφο σε Ομιλία (MP3) & Βίντεο (MP4)"
    page.window.width = 680
    page.window.height = 740
    page.theme_mode = ft.ThemeMode.DARK
    page.padding = 30
    page.window.center()

    # UI Elements
    selected_files_list = ft.ListView(expand=1, spacing=10, padding=20, auto_scroll=True)
    status_text = ft.Text("Κατάσταση: Σε αναμονή", size=16, color=ft.Colors.GREY_400)
    progress_bar = ft.ProgressBar(width=440, color="amber", bgcolor="#263238", value=0)
    progress_bar.visible = False

    file_queue = []

    def log(msg, error=False):
        color = ft.Colors.RED_400 if error else ft.Colors.GREEN_400
        selected_files_list.controls.append(ft.Text(msg, color=color))
        page.update()

    def set_all_buttons(disabled: bool):
        pick_btn.disabled = disabled
        clear_btn.disabled = disabled
        convert_btn.disabled = disabled
        video_btn.disabled = disabled
        page.update()

    async def pick_files_clicked(e):
        file_picker = ft.FilePicker()
        files_list = await file_picker.pick_files(allow_multiple=True, allowed_extensions=["pdf", "docx"])
        if files_list:
            for f in files_list:
                if f.path not in file_queue:
                    file_queue.append(f.path)
                    selected_files_list.controls.append(ft.Text(f"Επιλέχθηκε: {f.name}", color=ft.Colors.WHITE70))
            page.update()

    def clear_files(e):
        file_queue.clear()
        selected_files_list.controls.clear()
        selected_files_list.controls.append(ft.Text("Η λίστα καθαρίστηκε", color=ft.Colors.GREY_400))
        page.update()

    # ── MP3 Conversion ────────────────────────────────────────────────────────

    async def start_conversion(e):
        if not file_queue:
            status_text.value = "Παρακαλώ επιλέξτε αρχεία πρώτα!"
            status_text.color = ft.Colors.RED_400
            page.update()
            return

        set_all_buttons(True)
        progress_bar.visible = True
        progress_bar.value = 0
        page.update()

        for filepath in file_queue:
            status_text.value = f"Εξαγωγή κειμένου: {os.path.basename(filepath)}"
            status_text.color = ft.Colors.BLUE_400
            page.update()

            try:
                await asyncio.sleep(0.1)
                paragraphs = extract_paragraphs(filepath)
                if not paragraphs:
                    log(f"Δεν βρέθηκε κείμενο στο {os.path.basename(filepath)}", error=True)
                    continue

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
        status_text.color = ft.Colors.GREEN_400
        progress_bar.visible = False
        set_all_buttons(False)

    # ── MP4 Conversion ────────────────────────────────────────────────────────

    async def start_video_conversion(e):
        if not file_queue:
            status_text.value = "Παρακαλώ επιλέξτε αρχεία πρώτα!"
            status_text.color = ft.Colors.RED_400
            page.update()
            return

        set_all_buttons(True)
        progress_bar.visible = True
        progress_bar.value = 0
        page.update()

        for filepath in file_queue:
            ext = filepath.lower().split('.')[-1]
            if ext not in ('pdf', 'docx'):
                log(f"Μη υποστηριζόμενο αρχείο για βίντεο: {os.path.basename(filepath)}", error=True)
                continue

            status_text.value = f"🎬 Ξεκινά δημιουργία βίντεο: {os.path.basename(filepath)}"
            status_text.color = ft.Colors.PURPLE_300
            page.update()

            try:
                output_path = os.path.splitext(filepath)[0] + ".mp4"

                def update_video_progress(current, total, msg=""):
                    progress_bar.value = current / max(total, 1)
                    status_text.value = f"🎬 {msg}" if msg else f"🎬 Παράγραφος {current}/{total}"
                    status_text.color = ft.Colors.PURPLE_300
                    page.update()

                await convert_to_video(filepath, output_path, update_video_progress)
                log(f"🎬 Βίντεο: {os.path.basename(output_path)}")

            except Exception as ex:
                log(f"Σφάλμα βίντεο στο {os.path.basename(filepath)}: {str(ex)}", error=True)

        status_text.value = "Κατάσταση: Δημιουργία βίντεο ολοκληρώθηκε!"
        status_text.color = ft.Colors.GREEN_400
        progress_bar.visible = False
        set_all_buttons(False)

    # ── About dialog ──────────────────────────────────────────────────────────

    def open_about(e):
        dlg = ft.AlertDialog(
            title=ft.Text("🎵 Spyken", size=22, weight=ft.FontWeight.BOLD, color=ft.Colors.AMBER_400),
            content=ft.Column(
                [
                    ft.Text(f"Έκδοση: v{APP_VERSION}", size=15, color=ft.Colors.WHITE),
                    ft.Divider(height=10, color=ft.Colors.WHITE24),
                    ft.Text(
                        "Το Spyken μετατρέπει αρχεία .docx και .pdf σε αρχεία ήχου .mp3 "
                        "εναλλάσσοντας ανδρική και γυναικεία φωνή για κάθε παράγραφο, "
                        "καθώς και σε βίντεο .mp4 με οπτική ανάδειξη (highlight) κάθε παραγράφου.",
                        size=14, color=ft.Colors.GREY_300
                    ),
                    ft.Divider(height=10, color=ft.Colors.WHITE24),
                    ft.Text("🎶 Ελληνικές φωνές:", size=13, color=ft.Colors.BLUE_200),
                    ft.Text("  • Ανδρική: el-GR-NestorasNeural", size=12, color=ft.Colors.GREY_400),
                    ft.Text("  • Γυναικεία: el-GR-AthinaNeural", size=12, color=ft.Colors.GREY_400),
                    ft.Divider(height=5, color="transparent"),
                    ft.Text("🇬🇧 Αγγλικές φωνές (αυτόματη ανίχνευση):", size=13, color=ft.Colors.BLUE_200),
                    ft.Text("  • Ανδρική: en-GB-RyanNeural", size=12, color=ft.Colors.GREY_400),
                    ft.Text("  • Γυναικεία: en-GB-SoniaNeural", size=12, color=ft.Colors.GREY_400),
                    ft.Divider(height=10, color=ft.Colors.WHITE24),
                    ft.Text("by spyalekos • github.com/spyalekos", size=12, color=ft.Colors.GREY_500),
                ],
                tight=True,
            ),
            actions=[
                ft.TextButton("Κλείσιμο", on_click=lambda e: (setattr(dlg, 'open', False), page.update()))
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        page.overlay.append(dlg)
        dlg.open = True
        page.update()

    # ── Buttons ───────────────────────────────────────────────────────────────

    pick_btn = ft.ElevatedButton(
        "Επιλογή Αρχείων (.docx, .pdf)",
        icon=ft.Icons.FOLDER_OPEN,
        on_click=pick_files_clicked,
        style=ft.ButtonStyle(bgcolor=ft.Colors.INDIGO_700, color=ft.Colors.WHITE)
    )

    clear_btn = ft.ElevatedButton(
        "Καθαρισμός Λίστας",
        icon=ft.Icons.CLEAR,
        on_click=clear_files,
        style=ft.ButtonStyle(bgcolor=ft.Colors.RED_900, color=ft.Colors.WHITE)
    )

    convert_btn = ft.ElevatedButton(
        "Μετατροπή σε MP3",
        icon=ft.Icons.AUDIOTRACK,
        on_click=start_conversion,
        style=ft.ButtonStyle(bgcolor=ft.Colors.GREEN_700, color=ft.Colors.WHITE)
    )

    video_btn = ft.ElevatedButton(
        "Μετατροπή σε MP4",
        icon=ft.Icons.VIDEO_FILE,
        on_click=start_video_conversion,
        style=ft.ButtonStyle(bgcolor=ft.Colors.PURPLE_700, color=ft.Colors.WHITE)
    )

    # ── Layout ────────────────────────────────────────────────────────────────

    header = ft.Row(
        [
            ft.Icon(ft.Icons.AUDIOTRACK, size=40, color=ft.Colors.AMBER_400),
            ft.Text("Spyken", size=32, weight=ft.FontWeight.BOLD, color=ft.Colors.AMBER_400),
            ft.Text("by spyalekos  •  Έγγραφο σε MP3 & MP4", size=16, color=ft.Colors.GREY_300),
            ft.IconButton(
                icon=ft.Icons.INFO_OUTLINE,
                icon_color=ft.Colors.GREY_500,
                tooltip="Σχετικά",
                on_click=open_about,
            ),
        ],
        alignment=ft.MainAxisAlignment.CENTER,
    )

    button_row_top = ft.Row([pick_btn, clear_btn], alignment=ft.MainAxisAlignment.CENTER)
    button_row_bottom = ft.Row([convert_btn, video_btn], alignment=ft.MainAxisAlignment.CENTER, spacing=20)

    main_container = ft.Container(
        content=ft.Column(
            [
                header,
                ft.Divider(height=16, color="transparent"),
                button_row_top,
                ft.Container(
                    content=selected_files_list,
                    border=ft.border.all(1, ft.Colors.WHITE24),
                    border_radius=10,
                    padding=10,
                    expand=True
                ),
                ft.Divider(height=8, color="transparent"),
                button_row_bottom,
                ft.Divider(height=8, color="transparent"),
                ft.Column([status_text, progress_bar], horizontal_alignment=ft.CrossAxisAlignment.CENTER)
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        ),
        expand=True
    )

    page.add(main_container)


if __name__ == "__main__":
    ft.app(target=main)
