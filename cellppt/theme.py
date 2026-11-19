"""Colour palettes, fonts and slide geometry shared by every layout."""
from dataclasses import dataclass, field

from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches


def hex_rgb(value: str) -> RGBColor:
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


@dataclass
class Theme:
    name: str
    primary: str          # header bars, titles, cover band
    primary_dark: str
    accent: str           # **highlight** colour inside body text
    soft: str             # pale card header / chip background
    softer: str           # very pale card body
    text: str = "1F1F1F"
    muted: str = "6B6B6B"
    line: str = "D9D2E0"
    white: str = "FFFFFF"
    font_cn: str = "Microsoft YaHei"
    font_en: str = "Arial"
    font_title_cn: str = "Microsoft YaHei"
    gradient: list = field(default_factory=list)

    def rgb(self, key_or_hex: str) -> RGBColor:
        return hex_rgb(getattr(self, key_or_hex, key_or_hex))


THEMES = {
    # Matches the reference video: Nature-like aubergine purple + crimson highlights.
    "nature_purple": Theme("nature_purple", primary="6A1069", primary_dark="4E0B4D",
                           accent="C00000", soft="EADCF0", softer="F6F0F8",
                           gradient=["F3E8F5", "6A1069", "F3E8F5"]),
    "cell_blue": Theme("cell_blue", primary="0B4F8A", primary_dark="073763",
                       accent="D9480F", soft="DCEAF7", softer="F1F7FC",
                       gradient=["E6F0FA", "0B4F8A", "E6F0FA"]),
    "science_red": Theme("science_red", primary="9E1B1B", primary_dark="701313",
                         accent="1D5FA8", soft="F6DEDE", softer="FCF3F3",
                         gradient=["FAEAEA", "9E1B1B", "FAEAEA"]),
    "forest_green": Theme("forest_green", primary="1E6B4A", primary_dark="134A33",
                          accent="C2410C", soft="DDF0E6", softer="F2FAF5",
                          gradient=["E5F4EC", "1E6B4A", "E5F4EC"]),
}


def get_theme(name: str | None) -> Theme:
    return THEMES.get(name or "nature_purple", THEMES["nature_purple"])


# Slide geometry (16:9, 13.333in x 7.5in)
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.35)
HEADER_H = Inches(0.92)
CONTENT_TOP = Inches(1.08)


def emu(v) -> Emu:
    return Emu(int(v))
