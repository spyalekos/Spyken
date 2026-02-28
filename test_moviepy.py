import numpy as np
from moviepy import ImageClip, AudioArrayClip
from proglog import ProgressBarLogger

class TestLogger(ProgressBarLogger):
    def bars_callback(self, bar, attr, value, old_value=None):
        print(f"[{bar}] {attr} => {value}")

clip = ImageClip(np.zeros((10, 10, 3), dtype=np.uint8), duration=1.0)
audio = AudioArrayClip(np.zeros((44100, 2)), fps=44100)
clip = clip.with_audio(audio)
clip.write_videofile("test_out.mp4", fps=5, logger=TestLogger())
