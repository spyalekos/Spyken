import flet as ft
import flet_desktop
import pymupdf as fitz
import docx
import edge_tts
import asyncio
import os
import sys
import tempfile
import textwrap
import time

APP_VERSION = "1.6.9"

def get_asset_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        base_dir = os.path.abspath(".")
    p = os.path.join(base_dir, relative_path)
    if os.path.exists(p):
        return p
    return os.path.join(os.path.abspath("."), relative_path)

VOICE_MALE = "el-GR-NestorasNeural"
VOICE_FEMALE = "el-GR-AthinaNeural"
VOICE_EN_MALE = "en-GB-RyanNeural"
VOICE_EN_FEMALE = "en-GB-SoniaNeural"

# Voice speed rates (Female voices speak slightly faster by default, so we slow them down by 7% for balanced pacing)
RATE_MALE = "+0%"
RATE_FEMALE = "-7%"

def get_voice_rate(voice: str) -> str:
    """Returns the speech rate for the given voice."""
    if voice in (VOICE_FEMALE, VOICE_EN_FEMALE):
        return RATE_FEMALE
    return RATE_MALE

def resolve_voice(voice_str: str, text: str = "") -> str:
    """Resolves a user-friendly or localized voice name to an Edge-TTS voice identifier."""
    if not voice_str:
        return None
    v = voice_str.strip().lower()
    is_eng = is_english(text) if text else False
    if v in ("male", "man", "m", "ανδρας", "άνδρας", "ανδρικη", "ανδρική", "αντρικη", "αντρική"):
        return VOICE_EN_MALE if is_eng else VOICE_MALE
    if v in ("female", "woman", "f", "γυναικα", "γυναίκα", "γυναικεια", "γυναικεία"):
        return VOICE_EN_FEMALE if is_eng else VOICE_FEMALE
    return voice_str.strip()

def normalize_rate(rate_str: str) -> str:
    """Normalizes rate string to format like '+10%', '-7%'."""
    if not rate_str:
        return None
    r = rate_str.strip()
    if not r.endswith("%"):
        r = f"{r}%"
    if not r.startswith(("+", "-")):
        r = f"+{r}"
    return r

# Video frame resolution
VIDEO_W = 1280
VIDEO_H = 720
VIDEO_FPS = 10  # Reduced from 24 to speed up rendering significantly
VIDEO_TAIL_BUFFER_SEC = 1.5  # Freeze-frame buffer at the end of video for clear readability
PARAGRAPH_PAUSE_SEC = 0.4  # Natural breathing pause between paragraphs/speakers for smooth speech flow

# Marquee video frame resolution (horizontal banner, 1/10 of screen height)
MARQUEE_W = 1920
MARQUEE_H = 108
MARQUEE_FONT_SIZE = 50

# Fluorescent green highlight color
HIGHLIGHT_COLOR = (57, 255, 20)


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


import re as _re
_EMOJI_RE = _re.compile(
    "["
    "\U0001F000-\U0001FFFF"  # emoji, symbols, flags
    "\U00002300-\U000027BF"  # misc technical/dingbats
    "\U0000FE00-\U0000FE0F"  # variation selectors
    "\U0001FA00-\U0001FAFF"  # extended symbols
    "\U0000E000-\U0000F8FF"  # private use area
    "]+", flags=_re.UNICODE
)

def clean_document_text(text: str) -> str:
    """
    Cleans document text (especially DOCX) for visual rendering and speech:
    - Normalizes 'ό'τι' / 'Ό'τι' (and variants with curly/grave apostrophes) to 'ότι' / 'Ότι'
      so it is treated and pronounced as a single unified word instead of two words.
    - Strips all unreadable characters, emojis, symbol dingbats, and private use glyphs
      so that they do not appear as broken boxes in rendered video frames.
    - Strips soft hyphens, zero-width spaces, replacement characters, and control characters.
    - Normalizes multiple whitespace into a single space.
    """
    if not text:
        return ""
    # Normalize 'ό'τι' / 'Ό'τι' to 'ότι' / 'Ότι'
    text = _re.sub(r"(?:\b|^)[όΌ]['’`´]τι(?:\b|$)", lambda m: "Ότι" if m.group(0)[0].isupper() else "ότι", text)

    # Strip soft hyphens, zero-width spaces, joiners, BOM, and replacement char
    for ch in ('\xad', '\u200b', '\u200c', '\u200d', '\ufeff', '\ufffd'):
        text = text.replace(ch, '')
    text = text.replace('\xa0', ' ')

    # Strip emojis, symbols, and private-use area characters
    text = _EMOJI_RE.sub(' ', text)

    # Strip non-printable control characters
    text = ''.join(c if (ord(c) >= 32 or c in '\n\t') else ' ' for c in text)

    # Normalize excessive spaces
    text = _re.sub(r'\s+', ' ', text).strip()
    return text


def clean_for_tts(text: str) -> str:
    """
    Preprocess and clean text for natural speech synthesis without changing
    the sequence or token identities of words, ensuring 100% word-timing alignment
    with document text and word highlighting.
    """
    cleaned = clean_document_text(text)
    cleaned = _re.sub(r'\.{4,}', '... ', cleaned)
    cleaned = _re.sub(r'[-—–]{3,}', '— ', cleaned)
    cleaned = _re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned


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
            text = clean_document_text(p.text.strip())
            if is_valid_text(text):
                paragraphs.append(text)
    elif ext == 'pdf':
        doc = fitz.open(filepath)
        for page in doc:
            blocks = page.get_text("blocks")
            merged = merge_pdf_blocks(blocks)
            for (x0, y0, x1, y1, text) in merged:
                if is_valid_text(text):
                    paragraphs.append(text)
    else:
        raise ValueError("Μη υποστηριζόμενη μορφή αρχείου")

    return paragraphs


def merge_pdf_blocks(blocks: list) -> list:
    """
    Merge adjacent PDF text blocks that belong to the same paragraph.
    Two consecutive blocks are merged when they are vertically close
    (gap ≤ 1.8 × estimated line height) and horizontally overlapping.

    Input : list of fitz block tuples (x0, y0, x1, y1, text, block_no, block_type)
    Returns: list of (x0, y0, x1, y1, text) tuples.
    """
    # Keep only text blocks (block_type == 0)
    text_blocks = [b for b in blocks if b[6] == 0]

    if not text_blocks:
        return []

    merged = []
    cur_x0, cur_y0, cur_x1, cur_y1, cur_text = (
        text_blocks[0][0], text_blocks[0][1],
        text_blocks[0][2], text_blocks[0][3],
        text_blocks[0][4].strip(),
    )

    for b in text_blocks[1:]:
        bx0, by0, bx1, by1, btxt = b[0], b[1], b[2], b[3], b[4].strip()

        # Estimate line height of current block
        cur_height = cur_y1 - cur_y0
        lines_estimate = max(1, cur_text.count('\n') + 1)
        line_h = cur_height / lines_estimate

        # Vertical gap between end of current block and start of next
        v_gap = by0 - cur_y1

        # Horizontal overlap check: blocks share at least some horizontal range
        h_overlap = min(cur_x1, bx1) - max(cur_x0, bx0)

        if v_gap >= 0 and v_gap <= line_h * 1.8 and h_overlap > -20:
            # Merge: join with a space (replace internal newlines too)
            cur_text = cur_text + " " + btxt
            cur_x0 = min(cur_x0, bx0)
            cur_y1 = by1
            cur_x1 = max(cur_x1, bx1)
        else:
            merged.append((cur_x0, cur_y0, cur_x1, cur_y1, cur_text))
            cur_x0, cur_y0, cur_x1, cur_y1, cur_text = bx0, by0, bx1, by1, btxt

    merged.append((cur_x0, cur_y0, cur_x1, cur_y1, cur_text))

    # Clean up internal newlines in each merged block
    result = []
    for (x0, y0, x1, y1, t) in merged:
        clean = " ".join(t.splitlines()).strip()
        result.append((x0, y0, x1, y1, clean))

    return result


# ─────────────────────────── TTS WITH WORD TIMING ─────────────────────────────

def sync_edge_tts_clock():
    """
    Sync DRM clock skew with Microsoft's speech server date to prevent 403 Forbidden.
    Edge-TTS calculates the Sec-MS-GEC token using UTC time rounded to 5-minute buckets.
    If the system clock drifts even slightly across a 5-minute boundary, Microsoft rejects
    the WebSocket handshake with HTTP 403. This function fetches the server's HTTP Date
    header and synchronizes edge_tts.drm.DRM.clock_skew_seconds proactively.
    """
    try:
        import urllib.request
        from edge_tts.drm import DRM
        from datetime import datetime as dt, timezone as tz
        req = urllib.request.Request(
            "https://speech.platform.bing.com/consumer/speech/synthesize/readaloud/edge/v1",
            headers={"User-Agent": "Mozilla/5.0"}
        )
        urllib.request.urlopen(req, timeout=5)
    except Exception as e:
        headers = getattr(e, "headers", None)
        if headers:
            server_date = headers.get("Date")
            if server_date:
                from edge_tts.drm import DRM
                from datetime import datetime as dt, timezone as tz
                parsed = DRM.parse_rfc2616_date(server_date)
                if parsed:
                    local_now = dt.now(tz.utc).timestamp()
                    DRM.clock_skew_seconds = parsed - local_now


async def generate_tts_with_word_timings(text: str, voice: str, out_path: str, rate: str = None) -> list[dict]:
    """
    Stream TTS audio and collect WordBoundary events.
    Uses clean_for_tts(text) to strip emoji before sending to edge_tts,
    so word-boundary events contain proper words instead of character spans.
    Features resilient exponential backoff and clock resynchronization against rate-limiting.
    """
    tts_text = clean_for_tts(text)
    if not tts_text:
        return []
    if rate is None:
        rate = get_voice_rate(voice)
    word_timings = []
    audio_bytes = bytearray()

    delays = [1.0, 2.0, 3.5, 5.0, 7.0]
    for attempt, delay in enumerate(delays):
        try:
            communicate = edge_tts.Communicate(tts_text, voice, rate=rate, boundary="WordBoundary")
            word_timings.clear()
            audio_bytes.clear()

            async for chunk in communicate.stream():
                if chunk["type"] == "audio":
                    audio_bytes.extend(chunk["data"])
                elif chunk["type"] == "WordBoundary":
                    # offset and duration are in 100-nanosecond units
                    offset_s = chunk["offset"] / 1e7
                    duration_s = chunk["duration"] / 1e7
                    word = chunk.get("text", "")
                    word_timings.append({
                        "offset_s": offset_s,
                        "duration_s": duration_s,
                        "word": word,
                    })

            if audio_bytes and len(audio_bytes) > 500:
                with open(out_path, "wb") as f:
                    f.write(audio_bytes)
                return word_timings

        except Exception as e:
            err_str = str(e)
            if "403" in err_str or "Handshake" in err_str:
                sync_edge_tts_clock()

        await asyncio.sleep(delay)

    return []


async def generate_tts_chunk(text: str, voice: str, out_path: str, rate: str = None) -> bool:
    """Generate a single TTS mp3 chunk (audio only) with resilient backoff and clock resync."""
    tts_text = clean_for_tts(text)
    if not tts_text:
        return False
    if rate is None:
        rate = get_voice_rate(voice)
    delays = [1.0, 2.0, 3.5, 5.0, 7.0]
    for attempt, delay in enumerate(delays):
        try:
            communicate = edge_tts.Communicate(tts_text, voice, rate=rate)
            await communicate.save(out_path)
            if os.path.exists(out_path) and os.path.getsize(out_path) > 500:
                return True
        except Exception as e:
            err_str = str(e)
            if "403" in err_str or "Handshake" in err_str:
                sync_edge_tts_clock()
        await asyncio.sleep(delay)
    return False


# ─────────────────────────── AUDIO POST-PROCESSING ────────────────────────────

def get_ffmpeg_exe() -> str:
    """Locates the bundled or system ffmpeg binary executable."""
    try:
        import imageio_ffmpeg
        exe = imageio_ffmpeg.get_ffmpeg_exe()
        if exe and os.path.exists(exe):
            return exe
    except Exception:
        pass
    return "ffmpeg"


def post_process_audio_file(
    input_path: str,
    output_path: str = None,
    micro_fade: bool = True,
    warm_eq: bool = True,
    normalize_loudness: bool = True,
) -> bool:
    """
    Applies audio post-processing to synthetic speech audio:
    - Micro-fades (15ms in/out) to eliminate digital clicks/pops at boundaries.
    - High-pass filter (80 Hz) to eliminate low-frequency rumble and DC offset.
    - Vocal presence equalizer (3 kHz, +1.2 dB) for clarity and warmth.
    - Sibilance de-harshing (7.5 kHz, -2.0 dB) to soften harsh synthetic 's' sounds.
    - EBU R128 loudness normalization (-16 LUFS, True Peak -1.5 dB) for consistent volume.
    """
    import subprocess
    import shutil

    if output_path is None:
        output_path = input_path

    ffmpeg_exe = get_ffmpeg_exe()
    filters = []

    if micro_fade:
        filters.extend([
            "afade=t=in:ss=0:d=0.015",
            "areverse",
            "afade=t=in:ss=0:d=0.015",
            "areverse",
        ])
    if warm_eq:
        filters.extend([
            "highpass=f=80",
            "equalizer=f=3000:t=q:w=1.2:g=1.2",
            "equalizer=f=7500:t=q:w=1.5:g=-2.0",
        ])
    if normalize_loudness:
        filters.append("loudnorm=I=-16:TP=-1.5:LRA=11")

    filter_str = ",".join(filters)

    is_same = os.path.abspath(input_path) == os.path.abspath(output_path)
    target_out = (output_path + ".tmp_proc.mp3") if is_same else output_path

    try:
        cmd = [ffmpeg_exe, "-y", "-i", input_path]
        if filter_str:
            cmd.extend(["-af", filter_str])
        cmd.extend(["-b:a", "192k", target_out])

        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        res = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=creationflags,
        )
        if res.returncode == 0 and os.path.exists(target_out) and os.path.getsize(target_out) > 0:
            if is_same:
                shutil.move(target_out, output_path)
            return True
    except Exception:
        pass
    finally:
        if is_same and os.path.exists(target_out):
            try:
                os.remove(target_out)
            except Exception:
                pass

    if not is_same and input_path != output_path and os.path.exists(input_path):
        try:
            shutil.copy2(input_path, output_path)
            return True
        except Exception:
            pass

    return False


def create_silence_mp3(duration_sec: float, output_path: str) -> bool:
    """Generates an MP3 audio file containing clean silence of the specified duration."""
    import subprocess
    ffmpeg_exe = get_ffmpeg_exe()
    try:
        cmd = [
            ffmpeg_exe,
            "-y",
            "-f", "lavfi",
            "-i", "anullsrc=r=44100:cl=stereo",
            "-t", str(duration_sec),
            "-b:a", "192k",
            output_path,
        ]
        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        res = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=creationflags,
        )
        return res.returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0
    except Exception:
        return False


def concatenate_audio_files(input_files: list[str], output_path: str) -> bool:
    """Concatenates multiple MP3 audio files seamlessly using ffmpeg concat demuxer."""
    import subprocess
    if not input_files:
        return False

    ffmpeg_exe = get_ffmpeg_exe()
    list_file = output_path + ".concat_list.txt"
    try:
        with open(list_file, "w", encoding="utf-8") as f:
            for fpath in input_files:
                clean_p = os.path.abspath(fpath).replace("\\", "/")
                f.write(f"file '{clean_p}'\n")

        cmd = [
            ffmpeg_exe,
            "-y",
            "-f", "concat",
            "-safe", "0",
            "-i", list_file,
            "-c:a", "libmp3lame",
            "-b:a", "192k",
            output_path,
        ]
        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        res = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=creationflags,
        )
        return res.returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0
    except Exception:
        return False
    finally:
        if os.path.exists(list_file):
            try:
                os.remove(list_file)
            except Exception:
                pass


# ─────────────────────────── VIDEO HELPERS ────────────────────────────────────

def extract_paragraphs_pdf_with_pos(filepath: str) -> list[tuple]:
    """
    Returns list of (text, page_idx, fitz.Rect) for each valid paragraph.
    """
    results = []
    doc = fitz.open(filepath)
    for page_idx, page in enumerate(doc):
        blocks = page.get_text("blocks")
        merged = merge_pdf_blocks(blocks)
        for (x0, y0, x1, y1, text) in merged:
            if is_valid_text(text):
                rect = fitz.Rect(x0, y0, x1, y1)
                results.append((text, page_idx, rect, doc))
    return results


def get_pdf_word_rects(pdf_doc, page_idx: int, para_rect: fitz.Rect) -> list[tuple]:
    """
    Return list of (word_text, fitz.Rect) for words within para_rect on page_idx.
    Uses fitz word-level extraction.
    """
    page = pdf_doc[page_idx]
    # get_text("words") returns (x0, y0, x1, y1, "word", block_no, line_no, word_no)
    words = page.get_text("words")
    result = []
    for w in words:
        word_rect = fitz.Rect(w[0], w[1], w[2], w[3])
        word_text = w[4]
        # Skip words that are pure emoji / symbols (TTS won't produce word-boundaries for them)
        if not any(c.isalpha() or c.isdigit() for c in word_text):
            continue
        # Only include words that intersect with the paragraph block
        if word_rect.intersects(para_rect):
            result.append((word_text, word_rect))
    return result


def render_page_pdf_image(
    pdf_doc,
    page_idx: int,
    highlight_rect=None,
    word_highlight_rect=None,
    target_w: int = VIDEO_W,
    target_h: int = VIDEO_H,
):
    """
    Render a PDF page as a PIL RGBA image sized to fit target_w x target_h.
    - highlight_rect: paragraph-level semi-transparent highlight (fitz.Rect)
    - word_highlight_rect: word-level bright highlight (fitz.Rect), drawn on top
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

    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    def scaled_rect(r):
        return (
            int(r.x0 * scale) + x_off,
            int(r.y0 * scale) + y_off,
            int(r.x1 * scale) + x_off,
            int(r.y1 * scale) + y_off,
        )

    # Paragraph-level backdrop highlight (very faint)
    if highlight_rect is not None:
        hx0, hy0, hx1, hy1 = scaled_rect(highlight_rect)
        draw.rectangle([hx0, hy0, hx1, hy1], fill=(57, 255, 20, 30))
        draw.rectangle([hx0, hy0, hx1, hy1], outline=(57, 255, 20, 80), width=2)

    # Word-level highlight (bright on top)
    if word_highlight_rect is not None:
        wx0, wy0, wx1, wy1 = scaled_rect(word_highlight_rect)
        draw.rectangle([wx0, wy0, wx1, wy1], fill=(57, 255, 20, 120))
        draw.rectangle([wx0, wy0, wx1, wy1], outline=(57, 255, 20, 255), width=3)

    canvas = canvas.convert("RGBA")
    canvas = Image.alpha_composite(canvas, overlay).convert("RGB")

    return canvas


def _build_docx_layout(text: str, box_w: int, box_h: int):
    """
    Compute line wrapping and best font size for DOCX renderer.
    Returns (best_lines, best_font, line_h, font_size).
    """
    from PIL import ImageFont
    text = clean_document_text(text)

    def try_wrap(font_sz):
        try:
            font = ImageFont.truetype("arial.ttf", size=font_sz)
        except Exception:
            font = ImageFont.load_default()
            font_sz = 14

        words = text.split()
        if not words:
            return ["(κενή παράγραφος)"], font, font_sz

        lines = []
        current_line = []
        for word in words:
            test_line = " ".join(current_line + [word])
            w = font.getbbox(test_line)[2] if hasattr(font, 'getbbox') else font.getsize(test_line)[0]
            if w <= box_w:
                current_line.append(word)
            else:
                if current_line:
                    lines.append(" ".join(current_line))
                    current_line = [word]
                else:
                    lines.append(word)
                    current_line = []
        if current_line:
            lines.append(" ".join(current_line))

        return lines, font, font_sz

    best_lines = []
    best_font = None
    best_font_size = 50
    line_h = 50

    for f_size in range(50, 15, -2):
        lines, font, actual_f_size = try_wrap(f_size)
        test_line_h = max(22, int(actual_f_size * 1.5))
        total_h = len(lines) * test_line_h

        if total_h <= box_h or f_size == 16:
            best_lines = lines
            best_font = font
            best_font_size = actual_f_size
            line_h = test_line_h
            break

    if not best_lines:
        best_lines, best_font, best_font_size = try_wrap(24)
        line_h = 36

    return best_lines, best_font, line_h, best_font_size


_AVATAR_CACHE = {}

def get_avatar_image(name: str, target_h: int = 570):
    cache_key = (name, target_h)
    if cache_key not in _AVATAR_CACHE:
        from PIL import Image
        path = get_asset_path(os.path.join("assets", f"{name}.png"))
        if not os.path.exists(path):
            path = os.path.join("assets", f"{name}.png")
        if os.path.exists(path):
            try:
                raw_img = Image.open(path).convert("RGBA")
                av_w = int(raw_img.width * (target_h / raw_img.height))
                _AVATAR_CACHE[cache_key] = raw_img.resize((av_w, target_h), Image.Resampling.LANCZOS)
            except Exception:
                _AVATAR_CACHE[cache_key] = None
        else:
            _AVATAR_CACHE[cache_key] = None
    return _AVATAR_CACHE.get(cache_key)


def render_docx_paragraph_image(
    text: str,
    para_idx: int,
    total: int,
    highlight_word: str = None,
    highlight_word_idx: int = -1,
    mouth_state: int = -1,
    is_female: bool = False,
    target_w: int = VIDEO_W,
    target_h: int = VIDEO_H,
):
    """
    Render a DOCX paragraph as a PIL Image with a cartoon teacher avatar.
    - Avatar teacher stands on the left beside a presentation chalkboard.
    - Automatically switches between male and female teacher avatar matching speaker gender.
    - When words are spoken, the avatar animates its mouth (mouth_state).
    - If highlight_word is given, that specific word is highlighted in bright fluorescent green.
    Returns a PIL Image (RGB).
    """
    from PIL import Image, ImageDraw, ImageFont

    BG = (24, 28, 42)
    HEADER_BG = (40, 44, 64)
    H_COLOR = HIGHLIGHT_COLOR          # (57, 255, 20)
    TEXT_COLOR = (20, 20, 20)
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

    # Footer bar
    draw.rectangle([0, target_h - 40, target_w, target_h], fill=HEADER_BG)
    footer_txt = "Spyken · MP4 by spyalekos"
    try:
        footer_font = ImageFont.truetype("arial.ttf", size=18)
        draw.text((target_w // 2, target_h - 20), footer_txt,
                  font=footer_font, fill=(100, 100, 130), anchor="mm")
    except Exception:
        draw.text((target_w // 2, target_h - 20), footer_txt, fill=(100, 100, 130), anchor="mm")

    # Check for teacher avatar (male or female depending on is_female)
    av_h = 570
    prefix = "female_teacher" if is_female else "teacher"
    teacher_idle = get_avatar_image(f"{prefix}_idle", target_h=av_h)
    if teacher_idle is None and is_female:
        # Fallback to male teacher if female assets are missing
        teacher_idle = get_avatar_image("teacher_idle", target_h=av_h)
        prefix = "teacher"

    if teacher_idle is not None:
        # Choose avatar frame based on mouth_state (speaking vs resting)
        if mouth_state == 0:
            avatar = get_avatar_image(f"{prefix}_talk", target_h=av_h) or teacher_idle
        elif mouth_state == 1:
            avatar = get_avatar_image(f"{prefix}_talk2", target_h=av_h) or teacher_idle
        elif mouth_state == 2:
            avatar = teacher_idle
        else:
            avatar = teacher_idle

        # Presentation chalkboard on the right
        board_x = 410
        board_y = 80
        board_w = target_w - board_x - 30
        board_h = target_h - board_y - 65

        # Outer border / frame of board
        draw.rounded_rectangle(
            [board_x, board_y, board_x + board_w, board_y + board_h],
            radius=16, fill=(18, 22, 34), outline=(60, 70, 95), width=3
        )
        # Speech balloon pointer pointing left towards avatar mouth
        draw.polygon([(board_x, 320), (board_x - 35, 340), (board_x, 360)], fill=(18, 22, 34), outline=(60, 70, 95))
        draw.line([(board_x, 321), (board_x, 359)], fill=(18, 22, 34), width=3)

        # Inner board slate
        draw.rounded_rectangle(
            [board_x + 8, board_y + 8, board_x + board_w - 8, board_y + board_h - 8],
            radius=12, fill=(26, 32, 48)
        )
        draw.polygon([(board_x + 8, 325), (board_x - 25, 340), (board_x + 8, 355)], fill=(26, 32, 48))

        # Paste teacher avatar on the left side
        av_x = -70
        av_y = target_h - 40 - av_h
        canvas.paste(avatar, (av_x, av_y), avatar)

        # Text area layout inside board (generous margin ensuring no overlap with avatar hand)
        box_w = board_w - 110
        box_h = board_h - 80
        box_x = board_x + 75
        box_y = board_y + 40
        normal_txt_color = (215, 230, 245)
    else:
        # Fallback layout (box layout without avatar if assets unavailable)
        box_w = target_w - 120
        box_h = target_h - 160
        box_x = 60
        box_y = 90
        normal_txt_color = (200, 220, 200) if highlight_word else TEXT_COLOR

    best_lines, best_font, line_h, font_size = _build_docx_layout(text, box_w, box_h)

    # Compute overall text block dimensions
    text_block_h = len(best_lines) * line_h + 40
    max_line_w = 0
    for line in best_lines:
        w = best_font.getbbox(line)[2] if hasattr(best_font, 'getbbox') else best_font.getsize(line)[0]
        max_line_w = max(max_line_w, w)
    text_block_w = min(max_line_w + 40, box_w + 40)

    tx = box_x - 20
    ty = box_y - 10

    if teacher_idle is None:
        if highlight_word:
            draw.rectangle([tx, ty, tx + text_block_w, ty + text_block_h],
                           fill=(30, 80, 30))
            draw.rectangle([tx, ty, tx + text_block_w, ty + text_block_h],
                           outline=(0, 120, 0), width=2)
        else:
            draw.rectangle([tx, ty, tx + text_block_w, ty + text_block_h], fill=H_COLOR)
            draw.rectangle([tx, ty, tx + text_block_w, ty + text_block_h], outline=(0, 200, 0), width=4)

    # Draw text lines and find word bounding box
    y_cursor = box_y + 10

    def normalize(w):
        return w.strip(".,;:!?\"'()[]»«—–-").lower() if w else ""

    norm_hw = normalize(highlight_word)
    word_found = False
    current_word_idx = 0

    for line in best_lines:
        line_words = line.split()
        x_cursor = box_x + 10

        for lw in line_words:
            # Measure this word
            if hasattr(best_font, 'getbbox'):
                wb = best_font.getbbox(lw)
                w_width = wb[2] - wb[0]
                w_height = wb[3] - wb[1]
            else:
                w_width, w_height = best_font.getsize(lw)
                wb = (0, 0, w_width, w_height)

            # Draw the word
            should_highlight = False
            if highlight_word and not word_found:
                if highlight_word_idx != -1:
                    if current_word_idx == highlight_word_idx:
                        should_highlight = True
                else:
                    if normalize(lw) == norm_hw:
                        should_highlight = True

            if should_highlight:
                # Highlight this word
                pad = 4
                hx0 = x_cursor - pad
                hy0 = y_cursor - pad
                hx1 = x_cursor + w_width + pad
                hy1 = y_cursor + line_h - 2
                draw.rectangle([hx0, hy0, hx1, hy1], fill=H_COLOR)
                draw.rectangle([hx0, hy0, hx1, hy1], outline=(0, 220, 0), width=2)
                draw.text((x_cursor, y_cursor), lw, font=best_font, fill=TEXT_COLOR)
                word_found = True
            else:
                draw.text((x_cursor, y_cursor), lw, font=best_font, fill=normal_txt_color)

            # Advance x by word width + space width
            if hasattr(best_font, 'getbbox'):
                space_w = best_font.getbbox(" ")[2]
            else:
                space_w = best_font.getsize(" ")[0]
            x_cursor += w_width + space_w
            current_word_idx += 1

        y_cursor += line_h
        if y_cursor > box_y + box_h:
            break

    return canvas



# ─────────────────────── MARQUEE VIDEO HELPERS ────────────────────────────────

def prerender_marquee_canvas(paragraphs: list[str]):
    """
    Pre-render all paragraph text onto a wide virtual canvas for marquee video.
    Text is laid out horizontally in a single line that scrolls left-to-right.
    Returns (virtual_img, word_positions, font, total_canvas_width).
    word_positions: list of dicts {global_idx, x, y, w, h, word}.
    """
    from PIL import Image, ImageDraw, ImageFont

    padding_x = MARQUEE_W // 2  # half-viewport padding at start and end
    para_gap = 120              # horizontal gap between paragraphs

    try:
        font = ImageFont.truetype("arial.ttf", size=MARQUEE_FONT_SIZE)
    except Exception:
        font = ImageFont.load_default()

    if hasattr(font, 'getbbox'):
        space_w = font.getbbox(" ")[2]
    else:
        space_w = font.getsize(" ")[0]

    # Centre text vertically in the 108px strip
    text_y = (MARQUEE_H - MARQUEE_FONT_SIZE) // 2

    word_positions = []
    x_cursor = padding_x
    global_word_idx = 0

    for para in paragraphs:
        words = para.split()
        for word in words:
            if hasattr(font, 'getbbox'):
                wb = font.getbbox(word)
                word_w = wb[2] - wb[0]
            else:
                word_w = font.getsize(word)[0]

            word_positions.append({
                "global_idx": global_word_idx,
                "x": x_cursor,
                "y": text_y,
                "w": word_w,
                "h": MARQUEE_FONT_SIZE + 8,
                "word": word,
            })
            x_cursor += word_w + space_w
            global_word_idx += 1

        # Paragraph separator (extra horizontal space)
        x_cursor += para_gap

    total_canvas_width = x_cursor + padding_x

    virtual_img = Image.new("RGB", (total_canvas_width, MARQUEE_H), (20, 22, 30))
    draw = ImageDraw.Draw(virtual_img)

    for wp in word_positions:
        draw.text((wp["x"], wp["y"]), wp["word"], font=font, fill=(190, 200, 215))

    return virtual_img, word_positions, font, total_canvas_width


def get_mp3_duration(path: str) -> float:
    """Return duration of an mp3 file in seconds."""
    try:
        from mutagen.mp3 import MP3
        audio = MP3(path)
        return audio.info.length
    except Exception:
        return 3.0  # fallback


# ─────────────────────── WORD-TIMING ALIGNMENT ────────────────────────────────

def align_word_timings_to_text(word_timings: list[dict], text: str) -> list[dict]:
    """
    The TTS engine may return WordBoundary words in a slightly different order
    or with punctuation stripped. We do a best-effort match of timing words
    to the actual words in `text` so that highlights correspond to visible tokens.

    Returns same list but with 'text_word' and 'text_word_idx' keys containing the
    original word and its index from the text (or default/fallback if no match found).
    """
    text_words = text.split()
    timing_words = [t.copy() for t in word_timings]

    def _normalize(w: str) -> str:
        return (
            w.replace("\xad", "")
            .replace("\u200b", "")
            .replace("\ufeff", "")
            .strip(".,;:!?\"'()[]»«—–-")
            .lower()
        )

    t_idx = 0
    for tw in timing_words:
        # Advance text_words pointer to find the best match
        norm_tw = _normalize(tw["word"])
        matched = False
        for offset in range(min(12, len(text_words) - t_idx)):
            candidate = _normalize(text_words[t_idx + offset])
            if candidate == norm_tw or (
                norm_tw and candidate and (norm_tw in candidate or candidate in norm_tw)
            ):
                tw["text_word"] = text_words[t_idx + offset]
                tw["text_word_idx"] = t_idx + offset
                t_idx = t_idx + offset + 1
                matched = True
                break
        if not matched:
            tw["text_word"] = tw["word"]
            tw["text_word_idx"] = -1

    return timing_words


async def convert_to_video(filepath: str, output_path: str, progress_callback, voice_override: str = None, rate_override: str = None):
    """
    Master function: extract paragraphs (with positions for PDF),
    generate TTS per paragraph with word timings, build per-word video frames,
    assemble mp4.
    """
    import numpy as np
    from moviepy import ImageClip, AudioFileClip, VideoClip, concatenate_videoclips, concatenate_audioclips
    from proglog import ProgressBarLogger

    class FletMoviepyLogger(ProgressBarLogger):
        def __init__(self, cb):
            super().__init__()
            self.ui_callback = cb

        def bars_callback(self, bar, attr, value, old_value=None):
            if attr == 'index' and bar in ('chunk', 't', 'frame_index'):
                total = self.bars[bar].get('total', 1)
                if total > 0:
                    pct = int((value / total) * 100)
                    msg = "Μίξη ήχου..." if bar == 'chunk' else "Συναρμολόγηση βίντεο..."
                    self.ui_callback(value, total, f"{msg} {pct}%")

    ext = filepath.lower().split('.')[-1]
    temp_dir = tempfile.mkdtemp()

    try:
        # ── 1. Extract paragraphs ─────────────────────────────────────────────
        if ext == 'pdf':
            pdf_doc = fitz.open(filepath)
            para_data = []  # (text, page_idx, rect)
            for page_idx, page in enumerate(pdf_doc):
                blocks = page.get_text("blocks")
                merged = merge_pdf_blocks(blocks)
                for (x0, y0, x1, y1, text) in merged:
                    if is_valid_text(text):
                        rect = fitz.Rect(x0, y0, x1, y1)
                        para_data.append((text, page_idx, rect))
        else:  # docx
            doc_obj = docx.Document(filepath)
            para_data = []
            for p in doc_obj.paragraphs:
                text = clean_document_text(p.text.strip())
                if is_valid_text(text):
                    para_data.append((text, len(para_data), None))

        if not para_data:
            raise ValueError("Δεν βρέθηκαν παράγραφοι στο αρχείο.")

        total = len(para_data)
        clips = []
        voice_index = 0

        # ── 2. Per-paragraph: TTS with word timings + frames ──────────────────
        sync_edge_tts_clock()
        for i, para_item in enumerate(para_data):
            text = para_item[0]
            progress_callback(i, total, f"Παράγραφος {i+1}/{total}: TTS + λέξεις…")
            await asyncio.sleep(0.25)  # yield to UI and polite cooldown between requests

            # Pick voice
            if voice_override:
                voice = resolve_voice(voice_override, text)
            else:
                voice = (
                    (VOICE_EN_MALE if voice_index % 2 == 0 else VOICE_EN_FEMALE)
                    if is_english(text)
                    else (VOICE_MALE if voice_index % 2 == 0 else VOICE_FEMALE)
                )

            current_rate = rate_override if rate_override else get_voice_rate(voice)

            # For long paragraphs we chunk the text
            chunks = chunk_text(text, 800)
            chunk_audio_clips = []
            all_word_timings = []     # accumulated across chunks
            chunk_time_offset = 0.0  # running time offset for multi-chunk paragraphs

            for c_idx, chunk in enumerate(chunks):
                chunk_audio_path = os.path.join(temp_dir, f"audio_{i}_{c_idx}.mp3")
                word_timings = await generate_tts_with_word_timings(chunk, voice, chunk_audio_path, rate=current_rate)

                if os.path.exists(chunk_audio_path) and os.path.getsize(chunk_audio_path) > 500:
                    post_process_audio_file(chunk_audio_path, warm_eq=True, micro_fade=True, normalize_loudness=True)
                    ac = AudioFileClip(chunk_audio_path)
                    chunk_audio_clips.append(ac)

                    # Shift word timings by the running offset
                    for wt in word_timings:
                        shifted = dict(wt)
                        shifted["offset_s"] += chunk_time_offset
                        all_word_timings.append(shifted)

                    chunk_time_offset += ac.duration
                else:
                    # Fallback 1: plain TTS without timings
                    ok = await generate_tts_chunk(chunk, voice, chunk_audio_path, rate=current_rate)
                    if not ok or not os.path.exists(chunk_audio_path) or os.path.getsize(chunk_audio_path) <= 500:
                        # Fallback 2: try alternative voice (e.g. if one voice is rate-limited)
                        alt_voice = VOICE_MALE if voice == VOICE_FEMALE else (VOICE_FEMALE if voice == VOICE_MALE else voice)
                        alt_rate = rate_override if rate_override else get_voice_rate(alt_voice)
                        ok = await generate_tts_chunk(chunk, alt_voice, chunk_audio_path, rate=alt_rate)

                    if ok and os.path.exists(chunk_audio_path) and os.path.getsize(chunk_audio_path) > 500:
                        post_process_audio_file(chunk_audio_path, warm_eq=True, micro_fade=True, normalize_loudness=True)
                        ac = AudioFileClip(chunk_audio_path)
                        chunk_audio_clips.append(ac)
                        chunk_time_offset += ac.duration

            if chunk_audio_clips:
                if len(chunk_audio_clips) == 1:
                    combined_audio = chunk_audio_clips[0]
                else:
                    combined_audio = concatenate_audioclips(chunk_audio_clips)
                total_duration = combined_audio.duration
                voice_index += 1
            else:
                raise RuntimeError(f"Αποτυχία παραγωγής ήχου για την παράγραφο {i+1}. Παρακαλώ ελέγξτε τη σύνδεση στο διαδίκτυο.")

            # ── 3. Align word timings to actual text ──────────────────────────
            is_female_voice = voice in (VOICE_FEMALE, VOICE_EN_FEMALE)
            if all_word_timings:
                all_word_timings = align_word_timings_to_text(all_word_timings, text)
            elif ext == 'docx' and combined_audio is not None and total_duration > 0:
                # Resilient fallback for DOCX: synthesize proportional word timings
                # only when audio is present so word highlighting and talking mouth animation stay active!
                words = text.split()
                if words:
                    per_word_dur = total_duration / len(words)
                    all_word_timings = [
                        {
                            "offset_s": idx * per_word_dur,
                            "duration_s": per_word_dur,
                            "word": w,
                            "text_word": w,
                            "text_word_idx": idx,
                        }
                        for idx, w in enumerate(words)
                    ]

            # ── 4. Build per-word frames ───────────────────────────────────────
            progress_callback(i, total, f"Παράγραφος {i+1}/{total}: Frames…")
            await asyncio.sleep(0)

            # Pre-fetch PDF word rects (only for PDF with word timings)
            pdf_word_rects = []
            if ext == 'pdf' and all_word_timings:
                _, page_idx, para_rect = para_item
                pdf_word_rects = get_pdf_word_rects(pdf_doc, page_idx, para_rect)

            para_clips = []

            if all_word_timings:
                # ── Word-level frame generation ───────────────────────────────
                n_timings = len(all_word_timings)

                # For PDF: pre-render the base page image ONCE (no highlight),
                # then composite word highlights on top.
                # This avoids re-rendering the full page for every word.
                if ext == 'pdf':
                    _, page_idx, para_rect = para_item
                    base_pdf_img = render_page_pdf_image(
                        pdf_doc,
                        page_idx,
                        highlight_rect=para_rect,
                        word_highlight_rect=None,
                    )
                    # Normalise word rects for this page (scale + offset)
                    page = pdf_doc[page_idx]
                    page_rect = page.rect
                    scale = min(VIDEO_W / page_rect.width, VIDEO_H / page_rect.height)
                    x_off = (VIDEO_W - int(page_rect.width * scale)) // 2
                    y_off = (VIDEO_H - int(page_rect.height * scale)) // 2

                    # Sequential pointer into pdf_word_rects
                    wr_ptr = 0
                    word_rect_timings = []

                    for w_idx, wt in enumerate(all_word_timings):
                        if w_idx < n_timings - 1:
                            clip_dur = all_word_timings[w_idx + 1]["offset_s"] - wt["offset_s"]
                        else:
                            clip_dur = total_duration - wt["offset_s"]
                        if clip_dur < 0.04:
                            clip_dur = 0.04

                        word_text = wt.get("text_word", wt["word"])
                        norm_word = word_text.strip(".,;:!?\"'()[]\u00bb\u00ab\u2014\u2013-").lower()
                        matching_rect = None
                        search_limit = min(wr_ptr + 20, len(pdf_word_rects))
                        for j in range(wr_ptr, search_limit):
                            wr_text, wr_rect = pdf_word_rects[j]
                            wr_norm = wr_text.strip(".,;:!?\"'()[]\u00bb\u00ab\u2014\u2013-").lower()
                            if wr_norm == norm_word or (
                                wr_norm and norm_word and
                                (norm_word in wr_norm or wr_norm in norm_word)
                            ):
                                matching_rect = wr_rect
                                wr_ptr = j + 1
                                break
                        else:
                            # Do not artificially increment wr_ptr here to avoid cascade shifts
                            pass

                        word_rect_timings.append({
                            "start": wt["offset_s"],
                            "end": wt["offset_s"] + clip_dur,
                            "rect": matching_rect
                        })

                    def create_pdf_frame_func(wrt, base, s, xo, yo):
                        # Pre-convert base to numpy to avoid re-conversion in the loop
                        import numpy as np
                        from PIL import Image, ImageDraw
                        base_np = np.array(base.convert("RGB"))
                        # Cache for highlighted frames (key: rect as tuple)
                        cache = {None: base_np}

                        def make_frame(t):
                            active_rect = None
                            for entry in wrt:
                                if entry["start"] <= t < entry["end"]:
                                    active_rect = entry["rect"]
                                    break

                            # Create a hashable key for the cache
                            rect_key = (active_rect.x0, active_rect.y0, active_rect.x1, active_rect.y1) if active_rect else None
                            
                            if rect_key not in cache:
                                # Render this highlight ONCE and cache it
                                frame_img = base.copy()
                                overlay = Image.new("RGBA", (VIDEO_W, VIDEO_H), (0, 0, 0, 0))
                                draw = ImageDraw.Draw(overlay)
                                wx0 = int(active_rect.x0 * s) + xo
                                wy0 = int(active_rect.y0 * s) + yo
                                wx1 = int(active_rect.x1 * s) + xo
                                wy1 = int(active_rect.y1 * s) + yo
                                draw.rectangle([wx0, wy0, wx1, wy1], fill=(57, 255, 20, 140))
                                draw.rectangle([wx0, wy0, wx1, wy1], outline=(57, 255, 20, 255), width=3)
                                frame_img = Image.alpha_composite(frame_img.convert("RGBA"), overlay)
                                cache[rect_key] = np.array(frame_img.convert("RGB"))
                            
                            return cache[rect_key]
                        return make_frame

                    pdf_f = create_pdf_frame_func(word_rect_timings, base_pdf_img, scale, x_off, y_off)
                    para_clip_duration = (total_duration + VIDEO_TAIL_BUFFER_SEC) if (i == total - 1) else (total_duration + PARAGRAPH_PAUSE_SEC)
                    para_clips.append(VideoClip(pdf_f, duration=para_clip_duration))

                else:
                    # ── DOCX Lazy Rendering ───────────────────────────────────
                    _, para_idx, _ = para_item
                    word_render_timings = []
                    for w_idx, wt in enumerate(all_word_timings):
                        if w_idx < n_timings - 1:
                            clip_dur = all_word_timings[w_idx + 1]["offset_s"] - wt["offset_s"]
                        else:
                            clip_dur = total_duration - wt["offset_s"]
                        if clip_dur < 0.04:
                            clip_dur = 0.04

                        word_render_timings.append({
                            "start": wt["offset_s"],
                            "end": wt["offset_s"] + clip_dur,
                            "word": wt["word"],
                            "text_word_idx": wt.get("text_word_idx", -1)
                        })

                    def create_docx_frame_func(wrt, p_text, p_idx, p_total, female):
                        import numpy as np
                        cache = {}
                        def make_frame_docx(t):
                            active_word = None
                            active_word_idx = -1
                            for entry in wrt:
                                if entry["start"] <= t < entry["end"]:
                                    active_word = entry["word"]
                                    active_word_idx = entry["text_word_idx"]
                                    break
                            
                            # Animate avatar mouth when speaking, rest when silent
                            if active_word is not None:
                                mouth_state = int(t * 7) % 3
                            else:
                                mouth_state = -1

                            cache_key = (active_word, active_word_idx, mouth_state, female)
                            if cache_key not in cache:
                                img = render_docx_paragraph_image(
                                    p_text, p_idx, p_total,
                                    highlight_word=active_word,
                                    highlight_word_idx=active_word_idx,
                                    mouth_state=mouth_state,
                                    is_female=female
                                )
                                cache[cache_key] = np.array(img.convert("RGB"))
                            
                            return cache[cache_key]
                        return make_frame_docx

                    docx_f = create_docx_frame_func(word_render_timings, text, para_idx, total, is_female_voice)
                    para_clip_duration = (total_duration + VIDEO_TAIL_BUFFER_SEC) if (i == total - 1) else (total_duration + PARAGRAPH_PAUSE_SEC)
                    para_clips.append(VideoClip(docx_f, duration=para_clip_duration))

                # No gap-fill needed: offset-based durations already cover total_duration

            else:
                # ── Fallback: paragraph-level (original behaviour) ────────────
                if ext == 'pdf':
                    _, page_idx, rect = para_item
                    frame_img = render_page_pdf_image(pdf_doc, page_idx, highlight_rect=rect)
                else:
                    _, para_idx, _ = para_item
                    frame_img = render_docx_paragraph_image(text, para_idx, total, mouth_state=-1, is_female=is_female_voice)

                frame_np = np.array(frame_img.convert("RGB"))
                para_clip_duration = (total_duration + VIDEO_TAIL_BUFFER_SEC) if (i == total - 1) else (total_duration + PARAGRAPH_PAUSE_SEC)
                para_clips.append(ImageClip(frame_np, duration=para_clip_duration if para_clip_duration else 3.0))

            # ── 5. Concatenate word clips → paragraph clip ─────────────────────
            # (Note: para_clips now contains a single VideoClip for word-level, 
            # or a single ImageClip for fallback. Concatenation is mainly for multiple paragraphs below)
            if len(para_clips) == 1:
                para_video = para_clips[0]
            else:
                para_video = concatenate_videoclips(para_clips, method="chain")

            if combined_audio:
                if i == total - 1:
                    if VIDEO_TAIL_BUFFER_SEC > 0:
                        from moviepy import AudioArrayClip
                        silence = AudioArrayClip(np.zeros((int(44100 * VIDEO_TAIL_BUFFER_SEC), 2)), fps=44100)
                        combined_audio = concatenate_audioclips([combined_audio, silence])
                else:
                    if PARAGRAPH_PAUSE_SEC > 0:
                        from moviepy import AudioArrayClip
                        silence = AudioArrayClip(np.zeros((int(44100 * PARAGRAPH_PAUSE_SEC), 2)), fps=44100)
                        combined_audio = concatenate_audioclips([combined_audio, silence])
                audio_dur = combined_audio.duration
                video_dur = para_video.duration
                if audio_dur > video_dur:
                    combined_audio = combined_audio.subclipped(0, video_dur)
                para_video = para_video.with_audio(combined_audio)

            clips.append(para_video)

        # ── 6. Assemble final video ────────────────────────────────────────────
        progress_callback(total, total, "Συναρμολόγηση βίντεο...")
        await asyncio.sleep(0)

        ui_logger = FletMoviepyLogger(progress_callback)

        final = concatenate_videoclips(clips, method="chain")

        def _write_video():
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

        await asyncio.to_thread(_write_video)

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


async def convert_to_marquee_video(filepath: str, output_path: str, progress_callback, voice_override: str = None, rate_override: str = None):
    """
    Create a marquee-style video (MARQUEE_W × MARQUEE_H): text-only vertical
    scroll with word-level highlighting synchronized to TTS audio.
    Works for both PDF and DOCX inputs (text extraction only).
    """
    from moviepy import AudioFileClip, VideoClip, concatenate_audioclips
    from proglog import ProgressBarLogger

    class FletMoviepyLogger(ProgressBarLogger):
        def __init__(self, cb):
            super().__init__()
            self.ui_callback = cb

        def bars_callback(self, bar, attr, value, old_value=None):
            if attr == 'index' and bar in ('chunk', 't', 'frame_index'):
                total = self.bars[bar].get('total', 1)
                if total > 0:
                    pct = int((value / total) * 100)
                    msg = "Μίξη ήχου..." if bar == 'chunk' else "Marquee βίντεο..."
                    self.ui_callback(value, total, f"{msg} {pct}%")

    temp_dir = tempfile.mkdtemp()

    try:
        # 1. Extract paragraphs (text only, regardless of PDF or DOCX)
        paragraphs = extract_paragraphs(filepath)
        if not paragraphs:
            raise ValueError("Δεν βρέθηκαν παράγραφοι στο αρχείο.")

        total_paras = len(paragraphs)

        # 2. Pre-render marquee canvas
        progress_callback(0, total_paras, "Προετοιμασία marquee layout…")
        await asyncio.sleep(0)
        virtual_img, word_positions, font, canvas_width = prerender_marquee_canvas(paragraphs)

        if not word_positions:
            raise ValueError("Δεν βρέθηκαν λέξεις στο αρχείο.")

        # Build global_idx → position lookup
        wp_by_idx = {wp["global_idx"]: wp for wp in word_positions}

        # 3. TTS per paragraph with word timings
        sync_edge_tts_clock()
        all_audio_clips = []
        all_timed_words = []
        accumulated_time = 0.0
        global_word_offset = 0
        voice_index = 0

        for i, para_text in enumerate(paragraphs):
            progress_callback(i, total_paras, f"Παράγραφος {i+1}/{total_paras}: TTS…")
            await asyncio.sleep(0.25)

            if voice_override:
                voice = resolve_voice(voice_override, para_text)
            else:
                voice = (
                    (VOICE_EN_MALE if voice_index % 2 == 0 else VOICE_EN_FEMALE)
                    if is_english(para_text)
                    else (VOICE_MALE if voice_index % 2 == 0 else VOICE_FEMALE)
                )

            current_rate = rate_override if rate_override else get_voice_rate(voice)

            chunks = chunk_text(para_text, 800)
            chunk_audio_clips = []
            para_word_timings = []
            chunk_time_offset = 0.0

            for c_idx, chunk in enumerate(chunks):
                chunk_audio_path = os.path.join(temp_dir, f"mq_audio_{i}_{c_idx}.mp3")
                word_timings = await generate_tts_with_word_timings(
                    chunk, voice, chunk_audio_path, rate=current_rate
                )

                if os.path.exists(chunk_audio_path) and os.path.getsize(chunk_audio_path) > 500:
                    post_process_audio_file(chunk_audio_path, warm_eq=True, micro_fade=True, normalize_loudness=True)
                    ac = AudioFileClip(chunk_audio_path)
                    chunk_audio_clips.append(ac)
                    for wt in word_timings:
                        shifted = dict(wt)
                        shifted["offset_s"] += chunk_time_offset
                        para_word_timings.append(shifted)
                    chunk_time_offset += ac.duration
                else:
                    # Fallback 1: plain TTS without timings
                    ok = await generate_tts_chunk(chunk, voice, chunk_audio_path, rate=current_rate)
                    if not ok or not os.path.exists(chunk_audio_path) or os.path.getsize(chunk_audio_path) <= 500:
                        # Fallback 2: try alternative voice
                        alt_voice = VOICE_MALE if voice == VOICE_FEMALE else (VOICE_FEMALE if voice == VOICE_MALE else voice)
                        alt_rate = rate_override if rate_override else get_voice_rate(alt_voice)
                        ok = await generate_tts_chunk(chunk, alt_voice, chunk_audio_path, rate=alt_rate)

                    if ok and os.path.exists(chunk_audio_path) and os.path.getsize(chunk_audio_path) > 500:
                        post_process_audio_file(chunk_audio_path, warm_eq=True, micro_fade=True, normalize_loudness=True)
                        ac = AudioFileClip(chunk_audio_path)
                        chunk_audio_clips.append(ac)
                        chunk_time_offset += ac.duration

            if chunk_audio_clips:
                if len(chunk_audio_clips) == 1:
                    combined_audio = chunk_audio_clips[0]
                else:
                    combined_audio = concatenate_audioclips(chunk_audio_clips)
                para_duration = combined_audio.duration
                all_audio_clips.append(combined_audio)
                voice_index += 1
            else:
                raise RuntimeError(f"Αποτυχία παραγωγής ήχου για την παράγραφο {i+1} στο Marquee βίντεο.")

            # Align word timings and map to global word indices
            if para_word_timings:
                aligned = align_word_timings_to_text(para_word_timings, para_text)
                n_aligned = len(aligned)

                for w_idx, wt in enumerate(aligned):
                    if w_idx < n_aligned - 1:
                        clip_dur = aligned[w_idx + 1]["offset_s"] - wt["offset_s"]
                    else:
                        clip_dur = para_duration - wt["offset_s"]
                    if clip_dur < 0.04:
                        clip_dur = 0.04

                    local_idx = wt.get("text_word_idx", -1)
                    g_idx = (global_word_offset + local_idx) if local_idx >= 0 else -1

                    all_timed_words.append({
                        "start": accumulated_time + wt["offset_s"],
                        "end": accumulated_time + wt["offset_s"] + clip_dur,
                        "global_word_idx": g_idx,
                    })

            # Add inter-paragraph pause if not the last paragraph
            if i < total_paras - 1 and PARAGRAPH_PAUSE_SEC > 0:
                from moviepy import AudioArrayClip
                silence = AudioArrayClip(np.zeros((int(44100 * PARAGRAPH_PAUSE_SEC), 2)), fps=44100)
                all_audio_clips.append(silence)
                accumulated_time += para_duration + PARAGRAPH_PAUSE_SEC
            else:
                accumulated_time += para_duration

            global_word_offset += len(para_text.split())

        total_duration = accumulated_time
        if total_duration < 0.1:
            raise ValueError("Δεν παρήχθη ήχος.")

        # 4. Combine all audio clips + tail silence
        if VIDEO_TAIL_BUFFER_SEC > 0:
            from moviepy import AudioArrayClip
            silence = AudioArrayClip(np.zeros((int(44100 * VIDEO_TAIL_BUFFER_SEC), 2)), fps=44100)
            all_audio_clips.append(silence)

        if len(all_audio_clips) == 1:
            final_audio = all_audio_clips[0]
        else:
            final_audio = concatenate_audioclips(all_audio_clips)

        final_video_duration = total_duration + VIDEO_TAIL_BUFFER_SEC

        # 5. Build video with make_frame (scrolling viewport + word highlight)
        progress_callback(total_paras, total_paras, "Δημιουργία marquee frames…")
        await asyncio.sleep(0)

        def create_marquee_frame_func(timed_words, v_img, wp_idx, c_w, t_dur, fnt):
            import numpy as np
            from PIL import ImageDraw
            cache = {}

            def make_frame(t):
                active_g_idx = -1
                for entry in timed_words:
                    if entry["start"] <= t < entry["end"]:
                        active_g_idx = entry["global_word_idx"]
                        break

                cache_key = active_g_idx
                if cache_key in cache:
                    return cache[cache_key]

                # Determine horizontal scroll position based on active word
                if active_g_idx >= 0 and active_g_idx in wp_idx:
                    wp = wp_idx[active_g_idx]
                    target_x = wp["x"] + wp["w"] // 2
                else:
                    target_x = int(t / t_dur * (c_w - MARQUEE_W)) + MARQUEE_W // 2

                scroll_x = max(0, min(target_x - MARQUEE_W // 2, c_w - MARQUEE_W))

                # Crop horizontal viewport from virtual canvas
                viewport = v_img.crop(
                    (scroll_x, 0, scroll_x + MARQUEE_W, MARQUEE_H)
                ).copy()

                # Draw highlight on active word
                if active_g_idx >= 0 and active_g_idx in wp_idx:
                    wp = wp_idx[active_g_idx]
                    hl_x = wp["x"] - scroll_x
                    if -wp["w"] < hl_x < MARQUEE_W:
                        draw_vp = ImageDraw.Draw(viewport)
                        pad = 3
                        hx0 = max(0, hl_x - pad)
                        hy0 = max(0, wp["y"] - pad)
                        hx1 = min(MARQUEE_W, hl_x + wp["w"] + pad)
                        hy1 = min(MARQUEE_H, wp["y"] + wp["h"] + pad)
                        draw_vp.rectangle([hx0, hy0, hx1, hy1],
                                          fill=HIGHLIGHT_COLOR)
                        draw_vp.rectangle([hx0, hy0, hx1, hy1],
                                          outline=(0, 200, 0), width=2)
                        draw_vp.text((hl_x, wp["y"]), wp["word"],
                                     font=fnt, fill=(20, 22, 30))

                frame_np = np.array(viewport.convert("RGB"))
                cache[cache_key] = frame_np
                return frame_np

            return make_frame

        marquee_f = create_marquee_frame_func(
            all_timed_words, virtual_img, wp_by_idx,
            canvas_width, total_duration, font,
        )
        video_clip = VideoClip(marquee_f, duration=final_video_duration)
        video_clip = video_clip.with_audio(final_audio)

        # 6. Write final video
        ui_logger = FletMoviepyLogger(progress_callback)

        def _write_video():
            video_clip.write_videofile(
                output_path,
                fps=VIDEO_FPS,
                codec="libx264",
                audio_codec="aac",
                temp_audiofile=os.path.join(temp_dir, "tmp_mq_audio.m4a"),
                remove_temp=True,
                logger=ui_logger,
            )
            video_clip.close()

        await asyncio.to_thread(_write_video)

    finally:
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

async def convert_to_audio(paragraphs: list[str], output_path: str, progress_callback, voice_override: str = None, rate_override: str = None):
    """
    Generates seamless, natural MP3 speech audio from paragraphs:
    - Alternates male and female voices per paragraph.
    - Applies vocal presence EQ, high-pass filtering, and micro-fades to eliminate clicks.
    - Inserts a natural breathing pause (PARAGRAPH_PAUSE_SEC) between paragraphs.
    - Stitches audio files seamlessly via ffmpeg concat demuxer (with binary fallback).
    - Applies EBU R128 loudness normalization for consistent volume.
    """
    temp_dir = tempfile.mkdtemp()
    temp_files = []

    total_paras = len(paragraphs)
    voice_index = 0
    sync_edge_tts_clock()

    try:
        for p_idx, para in enumerate(paragraphs):
            await asyncio.sleep(0.25)
            if voice_override:
                voice = resolve_voice(voice_override, para)
            else:
                voice = (
                    (VOICE_EN_MALE if voice_index % 2 == 0 else VOICE_EN_FEMALE)
                    if is_english(para)
                    else (VOICE_MALE if voice_index % 2 == 0 else VOICE_FEMALE)
                )

            current_rate = rate_override if rate_override else get_voice_rate(voice)
            chunks = chunk_text(para, 800)
            para_success = False

            for c_idx, chunk in enumerate(chunks):
                temp_file = os.path.join(temp_dir, f"p{p_idx}_c{c_idx}.mp3")
                ok = await generate_tts_chunk(chunk, voice, temp_file, rate=current_rate)
                if not ok or not os.path.exists(temp_file) or os.path.getsize(temp_file) <= 500:
                    alt_voice = VOICE_MALE if voice == VOICE_FEMALE else (VOICE_FEMALE if voice == VOICE_MALE else voice)
                    alt_rate = rate_override if rate_override else get_voice_rate(alt_voice)
                    ok = await generate_tts_chunk(chunk, alt_voice, temp_file, rate=alt_rate)

                if ok and os.path.exists(temp_file) and os.path.getsize(temp_file) > 500:
                    post_process_audio_file(temp_file, warm_eq=True, micro_fade=True, normalize_loudness=False)
                    temp_files.append(temp_file)
                    para_success = True

            if para_success:
                voice_index += 1
                # Insert natural breathing pause between paragraphs (except after the last paragraph)
                if p_idx < total_paras - 1 and PARAGRAPH_PAUSE_SEC > 0:
                    silence_file = os.path.join(temp_dir, f"silence_{p_idx}.mp3")
                    if create_silence_mp3(PARAGRAPH_PAUSE_SEC, silence_file):
                        temp_files.append(silence_file)

            try:
                progress_callback(p_idx + 1, total_paras, f"Δημιουργία ήχου: {p_idx + 1}/{total_paras} παράγραφοι")
            except TypeError:
                progress_callback(p_idx + 1, total_paras)

        if not temp_files:
            raise ValueError("Δεν παρήχθησαν αρχεία ήχου.")

        # Stitch all parts and pauses together cleanly using ffmpeg concat
        try:
            progress_callback(total_paras, total_paras, "Ολοκλήρωση & ομαλοποίηση ήχου...")
        except TypeError:
            progress_callback(total_paras, total_paras)

        concat_ok = concatenate_audio_files(temp_files, output_path)
        if not concat_ok or not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            # Fallback to direct concatenation if ffmpeg fails
            with open(output_path, "wb") as outfile:
                for tf in temp_files:
                    with open(tf, "rb") as infile:
                        outfile.write(infile.read())

        # Final broadcast loudness normalization (EBU R128)
        post_process_audio_file(output_path, warm_eq=False, micro_fade=False, normalize_loudness=True)

    finally:
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
    page.title = "Spyken by spyalekos - Έγγραφο σε Ομιλία (MP3) & Βίντεο (MP4)"
    page.window.width = 680
    page.window.height = 740
    page.theme_mode = ft.ThemeMode.DARK
    page.padding = 30
    page.run_task(page.window.center)
    page.run_task(asyncio.to_thread, sync_edge_tts_clock)


    # UI Elements
    selected_files_list = ft.ListView(expand=1, spacing=10, padding=20, auto_scroll=True)
    status_text = ft.Text("Κατάσταση: Σε αναμονή", size=16, color=ft.Colors.GREY_400)
    progress_bar = ft.ProgressBar(width=440, color="amber", bgcolor="#263238", value=0)
    progress_bar.visible = False

    file_queue = []

    # ── Cartoon Teacher Avatar Card (DOCX only) ──────────────────────────────
    avatar_image = ft.Image(
        src=get_asset_path("assets/teacher_idle.png"),
        width=85,
        height=85,
        fit=ft.BoxFit.CONTAIN,
    )
    avatar_bubble_title = ft.Text("👨‍🏫 Δάσκαλος Spyken", weight=ft.FontWeight.BOLD, size=14, color=ft.Colors.AMBER_300)
    avatar_bubble_text = ft.Text("Επιλέξτε ένα αρχείο .docx για να ξεκινήσει η εκφώνηση!", size=13, color=ft.Colors.WHITE, italic=True)

    avatar_bubble = ft.Container(
        content=ft.Column(
            [
                ft.Row(
                    [
                        ft.Icon(ft.Icons.RECORD_VOICE_OVER, color=ft.Colors.AMBER_400, size=18),
                        avatar_bubble_title,
                    ],
                    spacing=6,
                ),
                ft.Divider(height=2, color="transparent"),
                avatar_bubble_text,
            ],
            spacing=2,
            tight=True,
        ),
        bgcolor="#1C2232",
        border=ft.Border.all(1, ft.Colors.BLUE_GREY_700),
        border_radius=12,
        padding=10,
        expand=True,
    )

    avatar_card = ft.Container(
        content=ft.Row(
            [
                ft.Container(
                    content=avatar_image,
                    border=ft.Border.all(2, ft.Colors.AMBER_400),
                    border_radius=45,
                    bgcolor="#181C2A",
                    padding=3,
                ),
                avatar_bubble,
            ],
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=12,
        ),
        bgcolor="#141824",
        border=ft.Border.all(1, ft.Colors.WHITE_24),
        border_radius=14,
        padding=8,
        visible=False,
    )

    _is_speaking = False

    async def avatar_speak(text: str, speed: float = 0.02, is_female: bool = False):
        nonlocal _is_speaking
        if not any(f.lower().endswith('.docx') for f in file_queue):
            return
        avatar_card.visible = True
        
        prefix = "female_teacher" if is_female else "teacher"

        if _is_speaking:
            avatar_bubble_text.value = text
            page.update()
            return

        _is_speaking = True
        try:
            mouth_images = [
                get_asset_path(f"assets/{prefix}_talk.png"),
                get_asset_path(f"assets/{prefix}_talk2.png"),
            ]
            idle_img = get_asset_path(f"assets/{prefix}_idle.png")

            step = 3
            for i in range(0, len(text), step):
                if not _is_speaking:
                    break
                avatar_bubble_text.value = text[:i + step]
                avatar_image.src = mouth_images[(i // step) % len(mouth_images)]
                page.update()
                await asyncio.sleep(speed)

            avatar_bubble_text.value = text
            avatar_image.src = idle_img
            page.update()
        except Exception:
            pass
        finally:
            _is_speaking = False

    def log(msg, error=False):
        color = ft.Colors.RED_400 if error else ft.Colors.GREEN_400
        selected_files_list.controls.append(ft.Text(msg, color=color))
        page.update()
        if any(f.lower().endswith('.docx') for f in file_queue):
            page.run_task(avatar_speak, msg)

    def set_all_buttons(disabled: bool):
        pick_btn.disabled = disabled
        clear_btn.disabled = disabled
        convert_btn.disabled = disabled
        video_btn.disabled = disabled
        marquee_btn.disabled = disabled
        page.update()

    async def pick_files_clicked(e):
        file_picker = ft.FilePicker()
        files_list = await file_picker.pick_files(allow_multiple=True, allowed_extensions=["pdf", "docx"])
        if files_list:
            docx_found = None
            for f in files_list:
                if f.path not in file_queue:
                    file_queue.append(f.path)
                    selected_files_list.controls.append(ft.Text(f"Επιλέχθηκε: {f.name}", color=ft.Colors.WHITE_70))
                    if f.name.lower().endswith('.docx'):
                        docx_found = f.name
            
            has_docx = any(p.lower().endswith('.docx') for p in file_queue)
            avatar_card.visible = has_docx
            page.update()

            if docx_found:
                await avatar_speak(f"Γεια σας! Φορτώθηκε το Word έγγραφο «{docx_found}». Είμαι έτοιμος να σας το διαβάσω!")

    def clear_files(e):
        file_queue.clear()
        selected_files_list.controls.clear()
        selected_files_list.controls.append(ft.Text("Η λίστα καθαρίστηκε", color=ft.Colors.GREY_400))
        avatar_card.visible = False
        avatar_image.src = get_asset_path("assets/teacher_idle.png")
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
            ext = filepath.lower().split('.')[-1]
            status_text.value = f"Εξαγωγή κειμένου: {os.path.basename(filepath)}"
            status_text.color = ft.Colors.BLUE_400
            page.update()

            if ext == 'docx':
                avatar_card.visible = True
                await avatar_speak(f"Ξεκινώ την ανάγνωση του «{os.path.basename(filepath)}»...")

            try:
                await asyncio.sleep(0.1)
                paragraphs = extract_paragraphs(filepath)
                if not paragraphs:
                    log(f"Δεν βρέθηκε κείμενο στο {os.path.basename(filepath)}", error=True)
                    continue

                output_path = os.path.splitext(filepath)[0] + ".mp3"

                def update_progress(current, total, msg=""):
                    progress_bar.value = (current / total) if total > 0 else 0
                    status_text.value = msg if msg else f"Δημιουργία ήχου: {current}/{total} παράγραφοι"
                    page.update()
                    if ext == 'docx' and current > 0:
                        fem = ((current - 1) % 2 != 0)
                        page.run_task(avatar_speak, f"Εκφωνώ παράγραφο {current}/{total}...", 0.02, fem)

                await convert_to_audio(paragraphs, output_path, update_progress)
                log(f"Ολοκληρώθηκε: {os.path.basename(output_path)}")

            except Exception as ex:
                log(f"Σφάλμα στο {os.path.basename(filepath)}: {str(ex)}", error=True)

        status_text.value = "Κατάσταση: Όλες οι μετατροπές ολοκληρώθηκαν!"
        status_text.color = ft.Colors.GREEN_400
        progress_bar.visible = False
        set_all_buttons(False)
        if any(p.lower().endswith('.docx') for p in file_queue):
            await avatar_speak("Όλες οι μετατροπές ολοκληρώθηκαν με επιτυχία!")

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

            if ext == 'docx':
                avatar_card.visible = True
                await avatar_speak(f"Ετοιμάζω το βίντεο παρουσίασης για το «{os.path.basename(filepath)}»...")

            try:
                output_path = os.path.splitext(filepath)[0] + ".mp4"
                loop = asyncio.get_running_loop()

                def update_video_progress(current, total, msg=""):
                    def _update_ui():
                        progress_bar.value = current / max(total, 1)
                        status_text.value = f"🎬 {msg}" if msg else f"🎬 Παράγραφος {current}/{total}"
                        status_text.color = ft.Colors.PURPLE_300
                        page.update()
                        if ext == 'docx' and current > 0:
                            fem = ((current - 1) % 2 != 0)
                            page.run_task(avatar_speak, f"Παράγω βίντεο παραγράφου {current}/{total}...", 0.02, fem)
                    loop.call_soon_threadsafe(_update_ui)

                await convert_to_video(filepath, output_path, update_video_progress)
                log(f"🎬 Βίντεο: {os.path.basename(output_path)}")

            except Exception as ex:
                log(f"Σφάλμα βίντεο στο {os.path.basename(filepath)}: {str(ex)}", error=True)

        status_text.value = "Κατάσταση: Δημιουργία βίντεο ολοκληρώθηκε!"
        status_text.color = ft.Colors.GREEN_400
        progress_bar.visible = False
        set_all_buttons(False)
        if any(p.lower().endswith('.docx') for p in file_queue):
            await avatar_speak("Η δημιουργία του βίντεο ολοκληρώθηκε με επιτυχία!")

    # ── Marquee Video Conversion ──────────────────────────────────────────────

    async def start_marquee_conversion(e):
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
                log(f"Μη υποστηριζόμενο αρχείο για marquee: {os.path.basename(filepath)}", error=True)
                continue

            status_text.value = f"📺 Ξεκινά δημιουργία Marquee video: {os.path.basename(filepath)}"
            status_text.color = ft.Colors.TEAL_300
            page.update()

            if ext == 'docx':
                avatar_card.visible = True
                await avatar_speak(f"Ετοιμάζω το Marquee banner για το «{os.path.basename(filepath)}»...")
            ext = filepath.lower().split('.')[-1]
            if ext not in ('pdf', 'docx'):
                log(f"Μη υποστηριζόμενο αρχείο για marquee: {os.path.basename(filepath)}", error=True)
                continue

            status_text.value = f"📜 Ξεκινά marquee βίντεο: {os.path.basename(filepath)}"
            status_text.color = ft.Colors.TEAL_300
            page.update()

            try:
                output_path = os.path.splitext(filepath)[0] + "_marquee.mp4"
                loop = asyncio.get_running_loop()

                def update_marquee_progress(current, total, msg=""):
                    def _update_ui():
                        progress_bar.value = current / max(total, 1)
                        status_text.value = f"📜 {msg}" if msg else f"📜 Παράγραφος {current}/{total}"
                        status_text.color = ft.Colors.TEAL_300
                        page.update()
                    loop.call_soon_threadsafe(_update_ui)

                await convert_to_marquee_video(filepath, output_path, update_marquee_progress)
                log(f"📜 Marquee: {os.path.basename(output_path)}")

            except Exception as ex:
                log(f"Σφάλμα marquee στο {os.path.basename(filepath)}: {str(ex)}", error=True)

        status_text.value = "Κατάσταση: Δημιουργία marquee ολοκληρώθηκε!"
        status_text.color = ft.Colors.GREEN_400
        progress_bar.visible = False
        set_all_buttons(False)

    # ── About dialog ──────────────────────────────────────────────────────────

    def open_about(e):
        dlg = ft.AlertDialog(
            title=ft.Text("🎵 Spyken by spyalekos", size=24, weight=ft.FontWeight.BOLD, color=ft.Colors.AMBER_400),
            content=ft.Container(
                content=ft.Column(
                    [
                        ft.Text(f"Έκδοση: v{APP_VERSION}", size=16, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                        ft.Divider(height=10, color=ft.Colors.WHITE_24),
                        
                        # 🌟 Δυνατότητες
                        ft.Text("🌟 Κύριες Δυνατότητες:", size=15, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_200),
                        ft.Text("• Μετατροπή σε MP3: Παραγωγή αρχείων ήχου με αυτόματη εναλλαγή ανδρικής & γυναικείας φωνής ανά παράγραφο.", size=13, color=ft.Colors.GREY_300),
                        ft.Text("• Ομαλή & Φυσική Ομιλία: Μετά-επεξεργασία ήχου (studio EQ, de-harshing, micro-fades, EBU R128 loudness normalization) και φυσικές ανάσες/παύσεις ανάμεσα στις παραγράφους.", size=13, color=ft.Colors.GREY_300),
                        ft.Text("• Μετατροπή σε MP4: Βίντεο παρουσίασης (1280×720) με οπτική προβολή των σελίδων PDF ή παραγράφων DOCX. Στα έγγραφα DOCX ενσωματώνεται διαδραστικός καρτούν-δάσκαλος avatar με δυναμική κίνηση ομιλίας συγχρονισμένη με την εκφώνηση και πράσινο φωτισμό κάθε λέξης.", size=13, color=ft.Colors.GREY_300),
                        ft.Text("• Marquee Video: Οριζόντιο banner (1920×108) με οριζόντια κυλιόμενο κείμενο και πράσινο φωσφορίζον highlight στην ενεργή λέξη. Ιδανικό για tickers και επικαλύψεις βίντεο.", size=13, color=ft.Colors.GREY_300),
                        ft.Text("• Έξυπνη Ανίχνευση Γλώσσας: Αυτόματη εναλλαγή μεταξύ Ελληνικών και Αγγλικών φωνών ανάλογα με το κείμενο.", size=13, color=ft.Colors.GREY_300),
                        ft.Text("• Πλήρης Γραμμή Εντολών (Headless CLI): Όλες οι λειτουργίες (εκφώνηση κειμένου, μετατροπή αρχείων σε MP3, MP4 ή Marquee Video) εκτελούνται απευθείας από το τερματικό χωρίς άνοιγμα του παραθύρου.", size=13, color=ft.Colors.GREY_300),
                        
                        ft.Divider(height=10, color=ft.Colors.WHITE_24),
                        
                        # 📖 Οδηγίες Χρήσης
                        ft.Text("📖 Οδηγίες Χρήσης:", size=15, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_200),
                        ft.Text("1. Επιλογή Αρχείων: Πατήστε το κουμπί για να επιλέξετε ένα ή περισσότερα αρχεία .docx ή .pdf.", size=13, color=ft.Colors.GREY_300),
                        ft.Text("2. Εκτέλεση Μετατροπής: Πατήστε ένα από τα τρία κουμπιά (MP3, MP4, Marquee Video) για να ξεκινήσει η επεξεργασία.", size=13, color=ft.Colors.GREY_300),
                        ft.Text("3. Αποθήκευση: Τα παραγόμενα αρχεία αποθηκεύονται αυτόματα στον ίδιο φάκελο με τα αρχικά σας έγγραφα με κατάλληλη ονοματολογία (_marquee.mp4, .mp4, .mp3).", size=13, color=ft.Colors.GREY_300),
                        ft.Text("4. Καθαρισμός: Πατήστε 'Καθαρισμός' αν θέλετε να αδειάσετε τη λίστα των επιλεγμένων αρχείων σας.", size=13, color=ft.Colors.GREY_300),
                        ft.Text("5. Χρήση από Τερματικό (CLI): Εκτελέστε 'Spyken.exe -h' ή 'Spyken.exe -v αρχείο.docx' απευθείας από τη γραμμή εντολών.", size=13, color=ft.Colors.GREY_300),
                        
                        ft.Divider(height=10, color=ft.Colors.WHITE_24),
                        
                        # 🎶 Φωνές
                        ft.Text("🎶 Φωνές που χρησιμοποιούνται:", size=14, color=ft.Colors.AMBER_200),
                        ft.Text("  • Ελληνική Ανδρική: el-GR-NestorasNeural", size=13, color=ft.Colors.GREY_400),
                        ft.Text("  • Ελληνική Γυναικεία: el-GR-AthinaNeural", size=13, color=ft.Colors.GREY_400),
                        ft.Text("  • Αγγλική Ανδρική: en-GB-RyanNeural", size=13, color=ft.Colors.GREY_400),
                        ft.Text("  • Αγγλική Γυναικεία: en-GB-SoniaNeural", size=13, color=ft.Colors.GREY_400),
                        
                        ft.Divider(height=10, color=ft.Colors.WHITE_24),
                        ft.Text("by spyalekos • github.com/spyalekos", size=13, color=ft.Colors.GREY_500),
                    ],
                    spacing=6,
                    scroll=ft.ScrollMode.AUTO,
                ),
                width=520,
                height=450,
                padding=10,
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

    pick_btn = ft.Button(
        "Επιλογή Αρχείων (.docx, .pdf)",
        icon=ft.Icons.FOLDER_OPEN,
        on_click=pick_files_clicked,
        style=ft.ButtonStyle(bgcolor=ft.Colors.INDIGO_700, color=ft.Colors.WHITE)
    )

    clear_btn = ft.Button(
        "Καθαρισμός Λίστας",
        icon=ft.Icons.CLEAR,
        on_click=clear_files,
        style=ft.ButtonStyle(bgcolor=ft.Colors.RED_900, color=ft.Colors.WHITE)
    )

    convert_btn = ft.Button(
        "Μετατροπή σε MP3",
        icon=ft.Icons.AUDIOTRACK,
        on_click=start_conversion,
        style=ft.ButtonStyle(bgcolor=ft.Colors.GREEN_700, color=ft.Colors.WHITE)
    )

    video_btn = ft.Button(
        "Μετατροπή σε MP4",
        icon=ft.Icons.VIDEO_FILE,
        on_click=start_video_conversion,
        style=ft.ButtonStyle(bgcolor=ft.Colors.PURPLE_700, color=ft.Colors.WHITE)
    )

    marquee_btn = ft.Button(
        "Marquee Video",
        icon=ft.Icons.SUBTITLES,
        on_click=start_marquee_conversion,
        style=ft.ButtonStyle(bgcolor=ft.Colors.TEAL_700, color=ft.Colors.WHITE)
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
    button_row_bottom = ft.Row([convert_btn, video_btn, marquee_btn], alignment=ft.MainAxisAlignment.CENTER, spacing=15)

    main_container = ft.Container(
        content=ft.Column(
            [
                header,
                avatar_card,
                ft.Divider(height=10, color="transparent"),
                button_row_top,
                ft.Container(
                    content=selected_files_list,
                    border=ft.Border.all(1, ft.Colors.WHITE_24),
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
    page.update()

    # Close PyInstaller Splash screen if running as compiled executable
    try:
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass


def play_audio_windows(file_path: str):
    import ctypes
    import time
    
    abs_path = os.path.abspath(file_path)
    alias = f"spyken_play_{os.getpid()}"
    open_cmd = f'open "{abs_path}" type mpegvideo alias {alias}'
    
    res = ctypes.windll.winmm.mciSendStringW(open_cmd, None, 0, 0)
    if res != 0:
        return False
        
    try:
        play_cmd = f'play {alias}'
        res = ctypes.windll.winmm.mciSendStringW(play_cmd, None, 0, 0)
        if res != 0:
            return False
            
        buffer = ctypes.create_unicode_buffer(128)
        while True:
            res = ctypes.windll.winmm.mciSendStringW(f'status {alias} mode', buffer, 128, 0)
            if res != 0:
                break
            mode = buffer.value.strip()
            if mode != 'playing':
                break
            time.sleep(0.05)
            
    finally:
        ctypes.windll.winmm.mciSendStringW(f'close {alias}', None, 0, 0)
        
    return True


def setup_windows_console():
    """Ensure standard output streams work cleanly in Windows console."""
    if sys.platform != "win32":
        return

    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    # If stdout/stderr are already piped or redirected (e.g. from PowerShell or CMD pipe)
    if sys.stdout is not None:
        try:
            if hasattr(sys.stdout, "reconfigure"):
                sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    if sys.stderr is not None:
        try:
            if hasattr(sys.stderr, "reconfigure"):
                sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    # If sys.stdout is None (standard PyInstaller GUI subsystem: console=False)
    if sys.stdout is None:
        ATTACH_PARENT_PROCESS = 0xFFFFFFFF
        kernel32.AttachConsole.argtypes = [wintypes.DWORD]
        kernel32.AttachConsole.restype = wintypes.BOOL

        attached = kernel32.AttachConsole(ATTACH_PARENT_PROCESS)
        if attached:
            try:
                kernel32.SetConsoleOutputCP(65001)
                kernel32.SetConsoleCP(65001)
            except Exception:
                pass
            try:
                sys.stdout = open("CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1)
                sys.stderr = open("CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1)
                sys.stdin = open("CONIN$", "r", encoding="utf-8", errors="replace")
            except Exception:
                pass

        if sys.stdout is None:
            sys.stdout = open(os.devnull, "w", encoding="utf-8")
        if sys.stderr is None:
            sys.stderr = open(os.devnull, "w", encoding="utf-8")


def create_temp_docx(text: str) -> str:
    """Creates a temporary .docx file with text paragraphs for video rendering."""
    temp_dir = tempfile.gettempdir()
    temp_docx = os.path.join(temp_dir, f"spyken_cli_text_{os.getpid()}_{int(time.time() * 1000)}.docx")
    doc = docx.Document()
    paragraphs = [p.strip() for p in text.splitlines() if p.strip()]
    if not paragraphs and text.strip():
        paragraphs = [text.strip()]
    for p in paragraphs:
        doc.add_paragraph(p)
    doc.save(temp_docx)
    return temp_docx


def cli_progress_callback(current, total, msg=""):
    """Command-line progress callback reporting percentage and current phase."""
    pct = int((current / max(total, 1)) * 100) if total > 0 else 0
    if msg:
        status_line = f"\r[{pct:3d}%] {msg}"
    else:
        status_line = f"\r[{pct:3d}%] Επεξεργασία {current}/{total}..."
    try:
        sys.stdout.write(status_line.ljust(75))
        sys.stdout.flush()
        if current >= total and total > 0 and ("100%" in msg or current == total):
            sys.stdout.write("\n")
            sys.stdout.flush()
    except Exception:
        pass


async def run_cli_async(text: str, voice_override: str = None, rate_override: str = None):
    """Plays text speech directly through speakers via Windows MCI API."""
    if voice_override:
        voice = resolve_voice(voice_override, text)
    else:
        voice = VOICE_EN_FEMALE if is_english(text) else VOICE_FEMALE
    rate = rate_override if rate_override else get_voice_rate(voice)
    print(f"Εκφώνηση: \"{text}\"")
    print(f"Φωνή: {voice} (Ρυθμός: {rate})")
    
    temp_dir = tempfile.gettempdir()
    temp_path = os.path.join(temp_dir, f"spyken_cli_temp_{os.getpid()}.mp3")
    
    try:
        print("Δημιουργία αρχείου ήχου...")
        ok = await generate_tts_chunk(text, voice, temp_path, rate=rate)
        if ok and os.path.exists(temp_path):
            post_process_audio_file(temp_path, warm_eq=True, micro_fade=True, normalize_loudness=True)
            print("Αναπαραγωγή...")
            play_audio_windows(temp_path)
            print("Ολοκληρώθηκε.")
        else:
            print("Σφάλμα: Αποτυχία δημιουργίας αρχείου ήχου μέσω Edge TTS.")
    except Exception as e:
        print(f"Σφάλμα κατά την εκτέλεση: {e}")
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def print_cli_help():
    """Prints comprehensive, beautifully formatted command line usage guide."""
    help_text = f"""Spyken v{APP_VERSION} - Μετατροπή Εγγράφων σε Ήχο (MP3) & Βίντεο (MP4)

ΧΡΗΣΗ:
  spyken [ΕΠΙΛΟΓΕΣ] [ΑΡΧΕΙΑ... | ΚΕΙΜΕΝΟ]

ΠΕΡΙΓΡΑΦΗ:
  Εκφωνεί κείμενο απευθείας από τα ηχεία ή μετατρέπει έγγραφα (.docx, .pdf)
  και απευθείας κείμενο σε αρχεία ήχου MP3, βίντεο παρουσίασης MP4 (με συγχρονισμένο
  καρτούν δάσκαλο/δασκάλα avatar και πράσινο φωτισμό λέξεων), ή οριζόντιο
  κυλιόμενο Marquee Banner MP4 (1920x108).
  Χωρίς ορίσματα εκκινεί το γραφικό περιβάλλον (GUI).

ΕΠΙΛΟΓΕΣ ΕΝΕΡΓΕΙΑΣ:
  -s, --speak                Εκφώνηση κειμένου απευθείας από τα ηχεία (προεπιλογή για κείμενο)
  -a, --audio, --mp3         Μετατροπή σε αρχείο ήχου MP3 (.mp3)
  -v, --video, --mp4         Μετατροπή σε βίντεο παρουσίασης MP4 (.mp4)
  -m, --marquee              Μετατροπή σε οριζόντιο κυλιόμενο Marquee Banner MP4 (.mp4)

ΕΠΙΛΟΓΕΣ ΕΙΣΟΔΟΥ & ΕΞΟΔΟΥ:
  -t, --text ΚΕΙΜΕΝΟ         Απευθείας εισαγωγή κειμένου για μετατροπή ή εκφώνηση
  -i, --input ΑΡΧΕΙΑ...      Ένα ή περισσότερα αρχεία εισόδου (.docx, .pdf)
  -o, --output ΔΙΑΔΡΟΜΗ      Όνομα αρχείου εξόδου ή φάκελος προορισμού

ΠΑΡΑΜΕΤΡΟΙ ΦΩΝΗΣ & ΗΧΟΥ:
  --voice ΦΩΝΗ               Επιλογή φωνής: 'male'/'ανδρική', 'female'/'γυναικεία',
                             ή όνομα φωνής Edge-TTS (π.χ. el-GR-NestorasNeural)
  -r, --rate ΡΥΘΜΟΣ          Ρύθμιση ταχύτητας ομιλίας (π.χ. +10%, -10%, +20%)

ΓΕΝΙΚΕΣ ΕΠΙΛΟΓΕΣ:
  -h, --help                 Εμφάνιση αυτού του μηνύματος βοήθειας
  -V, --version              Εμφάνιση της έκδοσης της εφαρμογής
  --gui                      Εκκίνηση του γραφικού περιβάλλοντος (GUI)

ΠΑΡΑΔΕΙΓΜΑΤΑ ΧΡΗΣΗΣ:
  • Απλή εκφώνηση κειμένου από τα ηχεία (υπάρχουσα λειτουργικότητα):
      Spyken.exe "Καλημέρα σας, καλή εβδομάδα!"
      Spyken.exe -s "Κείμενο προς εκφώνηση" --voice male --rate +10%

  • Μετατροπή εγγράφου σε ήχο MP3:
      Spyken.exe -a document.docx
      Spyken.exe --audio document.pdf -o "C:\\Audio\\podcast.mp3"

  • Μετατροπή εγγράφου σε βίντεο παρουσίασης MP4 (με avatar δασκάλου):
      Spyken.exe -v document.docx
      Spyken.exe --video presentation.docx -o presentation.mp4

  • Μετατροπή εγγράφου σε οριζόντιο Marquee Banner MP4 (1920x108):
      Spyken.exe -m document.docx
      Spyken.exe --marquee news.pdf -o news_ticker.mp4

  • Μετατροπή απευθείας κειμένου σε αρχείο:
      Spyken.exe -a -t "Κείμενο για αποθήκευση σε ήχο" -o audio.mp3
      Spyken.exe -v -t "Κείμενο για δημιουργία βίντεο" -o video.mp4

  • Μαζική μετατροπή πολλαπλών αρχείων:
      Spyken.exe -a file1.docx file2.pdf
      Spyken.exe -v *.docx
"""
    print(help_text)


async def cli_process_text(text: str, mode: str, output_path: str = None, voice: str = None, rate: str = None) -> int:
    """Processes plain text input from CLI into MP3 audio or MP4 video."""
    text = text.strip()
    if not text:
        print("Σφάλμα: Το κείμενο είναι κενό.")
        return 1

    if mode == "audio":
        if not output_path:
            output_path = "output.mp3"
        print(f"Spyken: Μετατροπή κειμένου σε ήχο MP3 -> {output_path}")
        paragraphs = [clean_document_text(p) for p in text.splitlines() if is_valid_text(clean_document_text(p))]
        if not paragraphs and is_valid_text(text):
            paragraphs = [clean_document_text(text)]
        if not paragraphs:
            print("Σφάλμα: Το κείμενο δεν περιέχει έγκυρες παραγράφους.")
            return 1
        try:
            await convert_to_audio(paragraphs, output_path, cli_progress_callback, voice_override=voice, rate_override=rate)
            print(f"\n[OK] Ολοκληρώθηκε με επιτυχία: {output_path}")
            return 0
        except Exception as e:
            print(f"\n[ΣΦΑΛΜΑ] {e}")
            return 1

    elif mode in ("video", "marquee"):
        default_ext = "_marquee.mp4" if mode == "marquee" else ".mp4"
        if not output_path:
            output_path = f"output{default_ext}"
        mode_label = "Marquee Video" if mode == "marquee" else "Βίντεο Παρουσίασης MP4"
        print(f"Spyken: Μετατροπή κειμένου σε {mode_label} -> {output_path}")

        temp_docx = create_temp_docx(text)
        try:
            if mode == "marquee":
                await convert_to_marquee_video(temp_docx, output_path, cli_progress_callback, voice_override=voice, rate_override=rate)
            else:
                await convert_to_video(temp_docx, output_path, cli_progress_callback, voice_override=voice, rate_override=rate)
            print(f"\n[OK] Ολοκληρώθηκε με επιτυχία: {output_path}")
            return 0
        except Exception as e:
            print(f"\n[ΣΦΑΛΜΑ] {e}")
            return 1
        finally:
            if os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except Exception:
                    pass
    return 0


async def cli_process_documents(files: list[str], mode: str, output_dest: str = None, voice: str = None, rate: str = None) -> int:
    """Processes document files (.docx, .pdf) in CLI mode."""
    valid_files = []
    for f in files:
        if not os.path.exists(f):
            print(f"[ΠΡΟΕΙΔΟΠΟΙΗΣΗ] Το αρχείο δεν βρέθηκε: {f}")
            continue
        ext = os.path.splitext(f)[1].lower()
        if ext not in (".docx", ".pdf"):
            print(f"[ΠΡΟΕΙΔΟΠΟΙΗΣΗ] Μη υποστηριζόμενος τύπος αρχείου (μόνο .docx και .pdf): {f}")
            continue
        valid_files.append(os.path.abspath(f))

    if not valid_files:
        print("Σφάλμα: Δεν βρέθηκαν έγκυρα αρχεία (.docx ή .pdf).")
        return 1

    total_files = len(valid_files)
    success_count = 0

    mode_label = {
        "audio": "Ήχο MP3",
        "video": "Βίντεο MP4",
        "marquee": "Marquee Video"
    }.get(mode, mode)

    print(f"\n=== Έναρξη επεξεργασίας {total_files} αρχείων σε {mode_label} ===")

    for idx, filepath in enumerate(valid_files, start=1):
        filename = os.path.basename(filepath)
        print(f"\n[{idx}/{total_files}] Επεξεργασία: {filename}")

        # Determine target output path
        if output_dest and total_files == 1 and not (os.path.isdir(output_dest) or output_dest.endswith(("/", "\\"))):
            target_out = os.path.abspath(output_dest)
        else:
            base_name, _ = os.path.splitext(filepath)
            suffix = ".mp3" if mode == "audio" else ("_marquee.mp4" if mode == "marquee" else ".mp4")
            out_filename = os.path.basename(base_name) + suffix
            if output_dest:
                os.makedirs(output_dest, exist_ok=True)
                target_out = os.path.join(os.path.abspath(output_dest), out_filename)
            else:
                target_out = os.path.join(os.path.dirname(filepath), out_filename)

        print(f"    Αρχείο εξόδου: {target_out}")

        try:
            if mode == "audio":
                paragraphs = extract_paragraphs(filepath)
                if not paragraphs:
                    print(f"    [ΣΦΑΛΜΑ] Δεν βρέθηκε κείμενο στο αρχείο {filename}")
                    continue
                await convert_to_audio(paragraphs, target_out, cli_progress_callback, voice_override=voice, rate_override=rate)
            elif mode == "video":
                await convert_to_video(filepath, target_out, cli_progress_callback, voice_override=voice, rate_override=rate)
            elif mode == "marquee":
                await convert_to_marquee_video(filepath, target_out, cli_progress_callback, voice_override=voice, rate_override=rate)

            if os.path.exists(target_out) and os.path.getsize(target_out) > 0:
                print(f"    [OK] Επιτυχής δημιουργία: {os.path.basename(target_out)} ({os.path.getsize(target_out):,} bytes)")
                success_count += 1
            else:
                print(f"    [ΣΦΑΛΜΑ] Το αρχείο εξόδου δεν δημιουργήθηκε σωστά.")
        except Exception as e:
            print(f"    [ΣΦΑΛΜΑ] {e}")

    print(f"\n=== Ολοκληρώθηκαν {success_count}/{total_files} αρχεία ===")
    return 0 if success_count > 0 else 1


def cli_main(raw_args: list[str]) -> int:
    """Master CLI dispatcher supporting both flags and direct text speech."""
    if not raw_args:
        return 0

    if any(a.lower() in ("-h", "--help", "/h", "/?", "-help", "--βοηθεια", "--βοήθεια") for a in raw_args):
        print_cli_help()
        return 0

    if any(a in ("-V", "--version", "-version", "/v", "/version") for a in raw_args if a not in ("-v",)):
        print(f"Spyken v{APP_VERSION}")
        return 0

    if "--gui" in raw_args:
        return -1

    # Check if any flags starting with '-' or '/' are present
    has_flags = any(a.startswith(("-", "/")) for a in raw_args)
    if not has_flags:
        # Check if all items are existing docx or pdf files
        all_are_docs = (
            len(raw_args) > 0 and
            all(os.path.isfile(a) and a.lower().endswith((".docx", ".pdf")) for a in raw_args)
        )
        if all_are_docs:
            print(f"Spyken v{APP_VERSION} - Μαζική μετατροπή σε ήχο MP3...")
            print("Σημείωση: Χρησιμοποιήστε -v για βίντεο MP4 ή -m για Marquee Video.")
            return asyncio.run(cli_process_documents(raw_args, mode="audio", output_dest=None, voice=None, rate=None))
        else:
            # Plain text speech playback (100% backward compatible!)
            text_to_speak = " ".join(raw_args).strip()
            if text_to_speak:
                asyncio.run(run_cli_async(text_to_speak))
            return 0

    # Parse arguments with argparse
    import argparse
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("-a", "--audio", "--mp3", action="store_true")
    parser.add_argument("-v", "--video", "--mp4", action="store_true")
    parser.add_argument("-m", "--marquee", action="store_true")
    parser.add_argument("-s", "--speak", action="store_true")
    parser.add_argument("-t", "--text", type=str, default=None)
    parser.add_argument("-i", "--input", nargs="*", default=[])
    parser.add_argument("-o", "--output", type=str, default=None)
    parser.add_argument("--voice", type=str, default=None)
    parser.add_argument("-r", "--rate", type=str, default=None)
    parser.add_argument("-V", "--version", action="store_true")
    parser.add_argument("-h", "--help", action="store_true")
    parser.add_argument("--gui", action="store_true")
    parser.add_argument("positional", nargs="*", default=[])

    try:
        parsed, remaining = parser.parse_known_args(raw_args)
    except Exception as e:
        print(f"Σφάλμα παραμέτρων: {e}")
        print("Εκτελέστε 'spyken --help' για αναλυτικές οδηγίες χρήσης.")
        return 1

    if parsed.help:
        print_cli_help()
        return 0

    if parsed.version:
        print(f"Spyken v{APP_VERSION}")
        return 0

    if parsed.gui:
        return -1

    voice_opt = resolve_voice(parsed.voice) if parsed.voice else None
    rate_opt = normalize_rate(parsed.rate) if parsed.rate else None

    # Collect inputs
    all_inputs = []
    if parsed.input:
        all_inputs.extend(parsed.input)
    if parsed.positional:
        all_inputs.extend(parsed.positional)
    if remaining:
        all_inputs.extend(remaining)

    # Resolve globbing for files
    import glob
    expanded_files = []
    for item in all_inputs:
        if any(c in item for c in ("*", "?")):
            matched = glob.glob(item)
            if matched:
                expanded_files.extend(matched)
            else:
                expanded_files.append(item)
        else:
            expanded_files.append(item)

    # Determine mode
    mode = None
    if parsed.audio:
        mode = "audio"
    elif parsed.video:
        mode = "video"
    elif parsed.marquee:
        mode = "marquee"
    elif parsed.speak:
        mode = "speak"

    text_input = parsed.text

    # If no mode explicitly set:
    if mode is None:
        if text_input:
            mode = "speak"
        elif expanded_files and all(os.path.isfile(f) and f.lower().endswith((".docx", ".pdf")) for f in expanded_files):
            mode = "audio"
        elif expanded_files:
            text_input = " ".join(expanded_files).strip()
            mode = "speak"
        else:
            print_cli_help()
            return 0

    # Execute determined mode
    if mode == "speak":
        if not text_input and expanded_files:
            text_input = " ".join(expanded_files).strip()
        if not text_input:
            print("Σφάλμα: Δεν δόθηκε κείμενο προς εκφώνηση.")
            return 1
        asyncio.run(run_cli_async(text_input, voice_override=voice_opt, rate_override=rate_opt))
        return 0

    if text_input and not expanded_files:
        return asyncio.run(cli_process_text(text_input, mode=mode, output_path=parsed.output, voice=voice_opt, rate=rate_opt))

    if not expanded_files:
        print("Σφάλμα: Δεν ορίστηκαν αρχεία εισόδου ή κείμενο.")
        print("Εκτελέστε 'spyken --help' για βοήθεια.")
        return 1

    return asyncio.run(cli_process_documents(expanded_files, mode=mode, output_dest=parsed.output, voice=voice_opt, rate=rate_opt))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Close PyInstaller Splash screen immediately if running from command line
        try:
            import pyi_splash
            pyi_splash.close()
        except ImportError:
            pass

        setup_windows_console()
        ret = cli_main(sys.argv[1:])
        if ret != -1:
            sys.exit(ret)

    ft.run(main, assets_dir=get_asset_path("assets"))
