#!/usr/bin/env python3

import os
from moviepy import TextClip, CompositeVideoClip, ColorClip

# User-configurable text
TITLE_TEXT = "古米特期货代采"
SUB_TEXT = "专业代采 一手货源"

# Output settings
WIDTH, HEIGHT = 1920, 1080
DURATION = 6.0
BG_COLOR = (0, 0, 0, 0)  # transparent
FPS = 30

# Font settings
FONT_FILE = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
TITLE_SIZE = 140
SUB_SIZE = 76
TITLE_COLOR = "white"
SUB_COLOR = "#FFD54F"


def build_clip():
    title = TextClip(
        font=FONT_FILE,
        text=TITLE_TEXT,
        font_size=TITLE_SIZE,
        color=TITLE_COLOR,
        size=(int(WIDTH*0.9), None),
        text_align="center",
        horizontal_align="center",
        vertical_align="center",
        transparent=True,
        duration=DURATION,
    )

    subtitle = TextClip(
        font=FONT_FILE,
        text=SUB_TEXT,
        font_size=SUB_SIZE,
        color=SUB_COLOR,
        size=(int(WIDTH*0.8), None),
        text_align="center",
        horizontal_align="center",
        vertical_align="center",
        transparent=True,
        duration=DURATION,
    )

    title = title.with_position(lambda t: ("center", HEIGHT*0.5 - 20 - 20*t))
    subtitle = subtitle.with_position(lambda t: ("center", HEIGHT*0.5 + 70 - 15*t))

    comp = CompositeVideoClip([title, subtitle], size=(WIDTH, HEIGHT))
    comp = comp.with_duration(DURATION).with_fps(FPS)
    return comp


def export_with_alpha_webm(clip: CompositeVideoClip, out_path: str):
    clip.write_videofile(
        out_path,
        codec="libvpx-vp9",
        fps=FPS,
        threads=os.cpu_count() or 4,
        preset="good",
        bitrate=None,
        audio=False,
        ffmpeg_params=[
            "-pix_fmt", "yuva420p",
            "-row-mt", "1",
            "-b:v", "0",
            "-crf", "28",
            "-deadline", "good",
            "-cpu-used", "4",
        ],
        temp_audiofile=None,
        remove_temp=True,
        write_logfile=False,
    )


def export_with_alpha_prores(clip: CompositeVideoClip, out_path: str):
    clip.write_videofile(
        out_path,
        codec="prores_ks",
        fps=FPS,
        threads=os.cpu_count() or 4,
        audio=False,
        ffmpeg_params=[
            "-profile:v", "4",  # 4444
            "-pix_fmt", "yuva444p10le",
            "-qscale:v", "9",
        ],
        temp_audiofile=None,
        remove_temp=True,
        write_logfile=False,
    )


def main():
    os.makedirs("/workspace/output", exist_ok=True)
    clip = build_clip()
    export_with_alpha_webm(clip, "/workspace/output/gumite_brand_anim.webm")
    export_with_alpha_prores(clip, "/workspace/output/gumite_brand_anim_prores4444.mov")


if __name__ == "__main__":
    main()