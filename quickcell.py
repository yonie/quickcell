#!/usr/bin/env python3
import os
import sys
from datetime import date, datetime, time

import gi

gi.require_version("Gtk", "3.0")
gi.require_version("Gdk", "3.0")
gi.require_version("Pango", "1.0")
gi.require_version("PangoCairo", "1.0")
from gi.repository import Gdk, GLib, Gtk, Pango, PangoCairo

import openpyxl
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils import get_column_letter

VERSION = "1.0.0"

DEFAULT_COL_WIDTH = 90
DEFAULT_ROW_HEIGHT = 22
HEADER_ROW_HEIGHT = 22
HEADER_COL_WIDTH = 52
CELL_PADDING = 4
MIN_COLS = 12
MIN_ROWS = 30
MIN_ZOOM = 0.3
MAX_ZOOM = 4.0

COLOR_GRID = (0.85, 0.85, 0.85)
COLOR_HEADER_BG = (0.94, 0.94, 0.94)
COLOR_HEADER_BG_ACTIVE = (0.82, 0.88, 0.96)
COLOR_HEADER_TEXT = (0.25, 0.25, 0.25)
COLOR_SEL_FILL = (0.28, 0.52, 0.85, 0.18)
COLOR_SEL_BORDER = (0.18, 0.42, 0.78)
COLOR_CELL_TEXT = (0.1, 0.1, 0.1)
COLOR_CELL_BG = (1.0, 1.0, 1.0)


# Default Office 2007+ theme palette (approximation — good enough for viewing)
_DEFAULT_THEME_RGB = [
    "FFFFFF",  # 0 bg1 / light1
    "000000",  # 1 tx1 / dark1
    "E7E6E6",  # 2 bg2 / light2
    "44546A",  # 3 tx2 / dark2
    "4472C4",  # 4 Accent1
    "ED7D31",  # 5 Accent2
    "A5A5A5",  # 6 Accent3
    "FFC000",  # 7 Accent4
    "5B9BD5",  # 8 Accent5
    "70AD47",  # 9 Accent6
    "0563C1",  # 10 Hlink
    "954F72",  # 11 FolHlink
]


def _parse_hex_rgb(s):
    if not s or not isinstance(s, str):
        return None
    s = s.strip()
    if len(s) == 8:
        s = s[2:]
    if len(s) != 6:
        return None
    try:
        return (
            int(s[0:2], 16) / 255,
            int(s[2:4], 16) / 255,
            int(s[4:6], 16) / 255,
        )
    except ValueError:
        return None


def _apply_tint(rgb, tint):
    r, g, b = rgb
    if tint < 0:
        f = 1 + tint
        return (max(0, r * f), max(0, g * f), max(0, b * f))
    f = tint
    return (r + (1 - r) * f, g + (1 - g) * f, b + (1 - b) * f)


def color_to_rgb(color):
    """Convert an openpyxl Color to an (r, g, b) tuple in [0,1], or None."""
    if color is None:
        return None
    t = getattr(color, "type", None)
    if t == "rgb":
        rgb = getattr(color, "rgb", None)
        if isinstance(rgb, str):
            return _parse_hex_rgb(rgb)
        return None
    if t == "indexed":
        idx = getattr(color, "indexed", None)
        if idx is None:
            return None
        try:
            return _parse_hex_rgb(COLOR_INDEX[idx])
        except (IndexError, TypeError):
            return None
    if t == "theme":
        idx = getattr(color, "theme", None)
        tint = getattr(color, "tint", 0) or 0
        if idx is None or not (0 <= idx < len(_DEFAULT_THEME_RGB)):
            return None
        base = _parse_hex_rgb(_DEFAULT_THEME_RGB[idx])
        if base is None:
            return None
        return _apply_tint(base, tint) if tint else base
    return None


def _fmt_stat(n):
    if n is None:
        return ""
    if isinstance(n, bool):
        return "TRUE" if n else "FALSE"
    try:
        f = float(n)
    except (TypeError, ValueError):
        return str(n)
    if f.is_integer() and abs(f) < 1e16:
        return str(int(f))
    s = f"{f:.6f}".rstrip("0").rstrip(".")
    return s or "0"


def _scroll_deltas(event):
    """Return (dx, dy) from a smooth-scroll event across PyGObject variants."""
    try:
        res = event.get_scroll_deltas()
    except Exception:
        return 0.0, 0.0
    if res is None:
        return 0.0, 0.0
    if len(res) == 3:
        _, dx, dy = res
    elif len(res) == 2:
        dx, dy = res
    else:
        return 0.0, 0.0
    return float(dx or 0.0), float(dy or 0.0)


def format_cell_value(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime):
        if value.hour or value.minute or value.second:
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, float):
        if value.is_integer() and abs(value) < 1e16:
            return str(int(value))
        return f"{value:.10g}"
    return str(value)


class SheetView(Gtk.Grid):
    def __init__(self, worksheet, on_selection_changed=None, on_zoom_changed=None):
        super().__init__()
        self.ws = worksheet
        self.on_selection_changed = on_selection_changed
        self.on_zoom_changed = on_zoom_changed

        self.zoom = 1.0

        self.max_row = max(worksheet.max_row or 0, MIN_ROWS)
        self.max_col = max(worksheet.max_column or 0, MIN_COLS)

        self.base_col_widths = []
        for i in range(1, self.max_col + 1):
            letter = get_column_letter(i)
            dim = worksheet.column_dimensions.get(letter)
            if dim and dim.width:
                px = int(dim.width * 7.0) + 5
                self.base_col_widths.append(max(20, px))
            else:
                self.base_col_widths.append(DEFAULT_COL_WIDTH)

        self.base_row_heights = []
        for i in range(1, self.max_row + 1):
            dim = worksheet.row_dimensions.get(i)
            if dim and dim.height:
                px = int(dim.height * 1.333)
                self.base_row_heights.append(max(14, px))
            else:
                self.base_row_heights.append(DEFAULT_ROW_HEIGHT)

        self.merged_anchor = {}
        self.merged_covered = {}
        for mr in worksheet.merged_cells.ranges:
            top, left = mr.min_row, mr.min_col
            bottom, right = mr.max_row, mr.max_col
            nrows = bottom - top + 1
            ncols = right - left + 1
            self.merged_anchor[(top, left)] = (nrows, ncols)
            for r in range(top, bottom + 1):
                for c in range(left, right + 1):
                    if (r, c) != (top, left):
                        self.merged_covered[(r, c)] = (top, left)

        self.sel_anchor = (1, 1)
        self.sel_cursor = (1, 1)
        self.dragging_sel = False
        self._drag_axis = "cell"

        self.scroll_x = 0
        self.scroll_y = 0

        self.drawing_area = Gtk.DrawingArea()
        self.drawing_area.set_can_focus(True)
        self.drawing_area.set_hexpand(True)
        self.drawing_area.set_vexpand(True)
        self.drawing_area.add_events(
            Gdk.EventMask.BUTTON_PRESS_MASK
            | Gdk.EventMask.BUTTON_RELEASE_MASK
            | Gdk.EventMask.POINTER_MOTION_MASK
            | Gdk.EventMask.SCROLL_MASK
            | Gdk.EventMask.SMOOTH_SCROLL_MASK
            | Gdk.EventMask.KEY_PRESS_MASK
            | Gdk.EventMask.FOCUS_CHANGE_MASK
        )
        self.drawing_area.connect("draw", self._on_draw)
        self.drawing_area.connect("button-press-event", self._on_button_press)
        self.drawing_area.connect("motion-notify-event", self._on_motion)
        self.drawing_area.connect("button-release-event", self._on_button_release)
        self.drawing_area.connect("scroll-event", self._on_scroll)
        self.drawing_area.connect("key-press-event", self._on_key_press)
        self.drawing_area.connect("size-allocate", self._on_size_allocate)

        self.hadj = Gtk.Adjustment(0, 0, 0, 20, 200, 0)
        self.vadj = Gtk.Adjustment(0, 0, 0, 20, 200, 0)
        self.hadj.connect("value-changed", self._on_hadj)
        self.vadj.connect("value-changed", self._on_vadj)
        self.hscroll = Gtk.Scrollbar(
            orientation=Gtk.Orientation.HORIZONTAL, adjustment=self.hadj
        )
        self.vscroll = Gtk.Scrollbar(
            orientation=Gtk.Orientation.VERTICAL, adjustment=self.vadj
        )

        self.attach(self.drawing_area, 0, 0, 1, 1)
        self.attach(self.vscroll, 1, 0, 1, 1)
        self.attach(self.hscroll, 0, 1, 1, 1)

    @property
    def col_widths(self):
        return [max(8, int(w * self.zoom)) for w in self.base_col_widths]

    @property
    def row_heights(self):
        return [max(8, int(h * self.zoom)) for h in self.base_row_heights]

    @property
    def header_row_height(self):
        return max(14, int(HEADER_ROW_HEIGHT * self.zoom))

    @property
    def header_col_width(self):
        return max(26, int(HEADER_COL_WIDTH * self.zoom))

    @property
    def content_width(self):
        return self.header_col_width + sum(self.col_widths)

    @property
    def content_height(self):
        return self.header_row_height + sum(self.row_heights)

    def _on_size_allocate(self, widget, rect):
        self._update_adjustments()

    def _update_adjustments(self):
        alloc = self.drawing_area.get_allocation()
        vw = max(1, alloc.width)
        vh = max(1, alloc.height)
        self.hadj.configure(
            self.hadj.get_value(), 0,
            max(vw, self.content_width), 20, max(1, int(vw * 0.8)), vw,
        )
        self.vadj.configure(
            self.vadj.get_value(), 0,
            max(vh, self.content_height), 20, max(1, int(vh * 0.8)), vh,
        )
        self.scroll_x = int(self.hadj.get_value())
        self.scroll_y = int(self.vadj.get_value())

    def _on_hadj(self, adj):
        self.scroll_x = int(adj.get_value())
        self.drawing_area.queue_draw()

    def _on_vadj(self, adj):
        self.scroll_y = int(adj.get_value())
        self.drawing_area.queue_draw()

    def _col_x(self, col_1based):
        x = self.header_col_width
        widths = self.col_widths
        for i in range(col_1based - 1):
            x += widths[i]
        return x

    def _row_y(self, row_1based):
        y = self.header_row_height
        heights = self.row_heights
        for i in range(row_1based - 1):
            y += heights[i]
        return y

    def _col_at_x(self, x_doc):
        if x_doc < self.header_col_width:
            return 0
        cx = self.header_col_width
        for i, w in enumerate(self.col_widths):
            if cx + w > x_doc:
                return i + 1
            cx += w
        return len(self.col_widths)

    def _row_at_y(self, y_doc):
        if y_doc < self.header_row_height:
            return 0
        cy = self.header_row_height
        for i, h in enumerate(self.row_heights):
            if cy + h > y_doc:
                return i + 1
            cy += h
        return len(self.row_heights)

    def _on_draw(self, widget, cr):
        alloc = widget.get_allocation()
        vw, vh = alloc.width, alloc.height

        cr.set_source_rgb(*COLOR_CELL_BG)
        cr.rectangle(0, 0, vw, vh)
        cr.fill()

        cws = self.col_widths
        rhs = self.row_heights
        hrh = self.header_row_height
        hcw = self.header_col_width

        first_col = last_col = None
        x = hcw
        for i, w in enumerate(cws):
            if x + w > self.scroll_x + hcw and x < self.scroll_x + vw:
                if first_col is None:
                    first_col = i + 1
                last_col = i + 1
            x += w
        first_row = last_row = None
        y = hrh
        for i, h in enumerate(rhs):
            if y + h > self.scroll_y + hrh and y < self.scroll_y + vh:
                if first_row is None:
                    first_row = i + 1
                last_row = i + 1
            y += h
        if first_col is None or first_row is None:
            return False

        header_font = Pango.FontDescription.new()
        header_font.set_family("Sans")
        header_font.set_size(int(9 * self.zoom * Pango.SCALE))

        sel_t, sel_l, sel_b, sel_r = self._sel_rect()

        cr.save()
        cr.rectangle(hcw, hrh, vw - hcw, vh - hrh)
        cr.clip()

        cr.set_source_rgb(*COLOR_GRID)
        cr.set_line_width(1)
        for row in range(first_row, last_row + 2):
            if row - 1 >= len(rhs):
                break
            y_doc = self._row_y(row)
            sy = y_doc - self.scroll_y
            cr.move_to(hcw, sy + 0.5)
            cr.line_to(vw, sy + 0.5)
            cr.stroke()
        for col in range(first_col, last_col + 2):
            if col - 1 >= len(cws):
                break
            x_doc = self._col_x(col)
            sx = x_doc - self.scroll_x
            cr.move_to(sx + 0.5, hrh)
            cr.line_to(sx + 0.5, vh)
            cr.stroke()

        for row in range(first_row, last_row + 1):
            for col in range(first_col, last_col + 1):
                if (row, col) in self.merged_covered:
                    continue
                anchor = self.merged_anchor.get((row, col))
                nrows, ncols = anchor if anchor else (1, 1)

                x_doc = self._col_x(col)
                y_doc = self._row_y(row)
                w_doc = sum(cws[col - 1 : col - 1 + ncols])
                h_doc = sum(rhs[row - 1 : row - 1 + nrows])
                sx = x_doc - self.scroll_x
                sy = y_doc - self.scroll_y

                cell = self.ws.cell(row=row, column=col)

                bg_rgb = None
                fill = cell.fill
                if fill is not None and getattr(fill, "patternType", None) == "solid":
                    bg_rgb = color_to_rgb(fill.fgColor) or color_to_rgb(
                        fill.start_color
                    )
                if bg_rgb is not None:
                    cr.set_source_rgb(*bg_rgb)
                    cr.rectangle(sx + 1, sy + 1, w_doc - 1, h_doc - 1)
                    cr.fill()
                elif nrows > 1 or ncols > 1:
                    cr.set_source_rgb(*COLOR_CELL_BG)
                    cr.rectangle(sx + 1, sy + 1, w_doc - 1, h_doc - 1)
                    cr.fill()

                if nrows > 1 or ncols > 1:
                    cr.set_source_rgb(*COLOR_GRID)
                    cr.set_line_width(1)
                    cr.rectangle(sx + 0.5, sy + 0.5, w_doc, h_doc)
                    cr.stroke()

                if sel_t <= row <= sel_b and sel_l <= col <= sel_r:
                    cr.set_source_rgba(*COLOR_SEL_FILL)
                    cr.rectangle(sx, sy, w_doc, h_doc)
                    cr.fill()

                text = format_cell_value(cell.value)
                if text:
                    font = cell.font
                    family = (font.name if font and font.name else "Sans")
                    size_pt = (font.size if font and font.size else 10) * self.zoom
                    bold = bool(font and font.bold)
                    italic = bool(font and font.italic)
                    fg_rgb = color_to_rgb(font.color) if font else None

                    font_desc = Pango.FontDescription.new()
                    font_desc.set_family(family)
                    font_desc.set_size(int(max(1, size_pt) * Pango.SCALE))
                    if bold:
                        font_desc.set_weight(Pango.Weight.BOLD)
                    if italic:
                        font_desc.set_style(Pango.Style.ITALIC)

                    cr.set_source_rgb(*(fg_rgb or COLOR_CELL_TEXT))
                    layout = PangoCairo.create_layout(cr)
                    layout.set_font_description(font_desc)
                    layout.set_text(text, -1)
                    avail = max(1, w_doc - 2 * CELL_PADDING)
                    layout.set_width(avail * Pango.SCALE)
                    layout.set_ellipsize(Pango.EllipsizeMode.END)

                    v = cell.value
                    align = getattr(cell.alignment, "horizontal", None)
                    if align == "left":
                        pango_align = Pango.Alignment.LEFT
                    elif align == "right":
                        pango_align = Pango.Alignment.RIGHT
                    elif align in ("center", "centerContinuous"):
                        pango_align = Pango.Alignment.CENTER
                    elif isinstance(v, bool):
                        pango_align = Pango.Alignment.CENTER
                    elif isinstance(v, (int, float, datetime, date, time)):
                        pango_align = Pango.Alignment.RIGHT
                    else:
                        pango_align = Pango.Alignment.LEFT
                    layout.set_alignment(pango_align)

                    _, logical = layout.get_pixel_extents()
                    ty = sy + max(CELL_PADDING, (h_doc - logical.height) / 2)
                    cr.move_to(sx + CELL_PADDING, ty)
                    PangoCairo.show_layout(cr, layout)

        self._stroke_selection_border(cr, font_desc, sel_t, sel_l, sel_b, sel_r)
        cr.restore()

        cr.save()
        cr.rectangle(hcw, 0, vw - hcw, hrh)
        cr.clip()
        for col in range(first_col, last_col + 1):
            x_doc = self._col_x(col)
            w = cws[col - 1]
            sx = x_doc - self.scroll_x
            active = sel_l <= col <= sel_r
            bg = COLOR_HEADER_BG_ACTIVE if active else COLOR_HEADER_BG
            cr.set_source_rgb(*bg)
            cr.rectangle(sx, 0, w, hrh)
            cr.fill()
            cr.set_source_rgb(*COLOR_GRID)
            cr.set_line_width(1)
            cr.rectangle(sx + 0.5, 0.5, w, hrh)
            cr.stroke()
            cr.set_source_rgb(*COLOR_HEADER_TEXT)
            layout = PangoCairo.create_layout(cr)
            layout.set_font_description(header_font)
            layout.set_text(get_column_letter(col), -1)
            layout.set_width(w * Pango.SCALE)
            layout.set_alignment(Pango.Alignment.CENTER)
            _, logical = layout.get_pixel_extents()
            cr.move_to(sx, (hrh - logical.height) / 2)
            PangoCairo.show_layout(cr, layout)
        cr.restore()

        cr.save()
        cr.rectangle(0, hrh, hcw, vh - hrh)
        cr.clip()
        for row in range(first_row, last_row + 1):
            y_doc = self._row_y(row)
            h = rhs[row - 1]
            sy = y_doc - self.scroll_y
            active = sel_t <= row <= sel_b
            bg = COLOR_HEADER_BG_ACTIVE if active else COLOR_HEADER_BG
            cr.set_source_rgb(*bg)
            cr.rectangle(0, sy, hcw, h)
            cr.fill()
            cr.set_source_rgb(*COLOR_GRID)
            cr.set_line_width(1)
            cr.rectangle(0.5, sy + 0.5, hcw, h)
            cr.stroke()
            cr.set_source_rgb(*COLOR_HEADER_TEXT)
            layout = PangoCairo.create_layout(cr)
            layout.set_font_description(header_font)
            layout.set_text(str(row), -1)
            layout.set_width(hcw * Pango.SCALE)
            layout.set_alignment(Pango.Alignment.CENTER)
            _, logical = layout.get_pixel_extents()
            cr.move_to(0, sy + (h - logical.height) / 2)
            PangoCairo.show_layout(cr, layout)
        cr.restore()

        cr.set_source_rgb(*COLOR_HEADER_BG)
        cr.rectangle(0, 0, hcw, hrh)
        cr.fill()
        cr.set_source_rgb(*COLOR_GRID)
        cr.set_line_width(1)
        cr.rectangle(0.5, 0.5, hcw, hrh)
        cr.stroke()
        return False

    def _stroke_selection_border(self, cr, font_desc, sel_t, sel_l, sel_b, sel_r):
        rb, cb = sel_b, sel_r
        for (r, c), (nr, nc) in self.merged_anchor.items():
            if sel_t <= r <= sel_b and sel_l <= c <= sel_r:
                rb = max(rb, r + nr - 1)
                cb = max(cb, c + nc - 1)
        x_doc = self._col_x(sel_l)
        y_doc = self._row_y(sel_t)
        right_doc = self._col_x(cb + 1)
        bottom_doc = self._row_y(rb + 1)
        sx = x_doc - self.scroll_x
        sy = y_doc - self.scroll_y
        sw = right_doc - x_doc
        sh = bottom_doc - y_doc
        cr.set_source_rgb(*COLOR_SEL_BORDER)
        cr.set_line_width(2)
        cr.rectangle(sx + 0.5, sy + 0.5, sw - 1, sh - 1)
        cr.stroke()

    def _sel_rect(self):
        r1, c1 = self.sel_anchor
        r2, c2 = self.sel_cursor
        return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)

    def _on_button_press(self, widget, event):
        self.drawing_area.grab_focus()
        if event.button != 1:
            return False
        x_doc = event.x + self.scroll_x
        y_doc = event.y + self.scroll_y
        col = self._col_at_x(x_doc)
        row = self._row_at_y(y_doc)

        if row == 0 and col == 0:
            self.sel_anchor = (1, 1)
            self.sel_cursor = (self.max_row, self.max_col)
            self.dragging_sel = False
        elif row == 0:
            self.sel_anchor = (1, col)
            self.sel_cursor = (self.max_row, col)
            self.dragging_sel = True
            self._drag_axis = "col"
        elif col == 0:
            self.sel_anchor = (row, 1)
            self.sel_cursor = (row, self.max_col)
            self.dragging_sel = True
            self._drag_axis = "row"
        else:
            if event.state & Gdk.ModifierType.SHIFT_MASK:
                self.sel_cursor = (row, col)
            else:
                self.sel_anchor = (row, col)
                self.sel_cursor = (row, col)
            self.dragging_sel = True
            self._drag_axis = "cell"

        self._notify_selection()
        self.drawing_area.queue_draw()
        return True

    def _on_motion(self, widget, event):
        if not self.dragging_sel:
            return False
        x_doc = event.x + self.scroll_x
        y_doc = event.y + self.scroll_y
        col = max(1, min(self.max_col, self._col_at_x(x_doc) or 1))
        row = max(1, min(self.max_row, self._row_at_y(y_doc) or 1))
        if self._drag_axis == "col":
            self.sel_cursor = (self.max_row, col)
        elif self._drag_axis == "row":
            self.sel_cursor = (row, self.max_col)
        else:
            self.sel_cursor = (row, col)
        self._notify_selection()
        self.drawing_area.queue_draw()
        return True

    def _on_button_release(self, widget, event):
        self.dragging_sel = False
        return True

    def _on_scroll(self, widget, event):
        if event.state & Gdk.ModifierType.CONTROL_MASK:
            old_zoom = self.zoom
            if event.direction == Gdk.ScrollDirection.SMOOTH:
                _, dy = _scroll_deltas(event)
                if dy < 0:
                    self.zoom *= 1.1 ** abs(dy)
                elif dy > 0:
                    self.zoom /= 1.1 ** abs(dy)
            elif event.direction == Gdk.ScrollDirection.UP:
                self.zoom *= 1.1
            elif event.direction == Gdk.ScrollDirection.DOWN:
                self.zoom /= 1.1
            else:
                return False
            self.zoom = max(MIN_ZOOM, min(MAX_ZOOM, self.zoom))
            if abs(self.zoom - old_zoom) < 1e-6:
                return True

            ratio = self.zoom / old_zoom
            self.scroll_x = int((self.scroll_x + event.x) * ratio - event.x)
            self.scroll_y = int((self.scroll_y + event.y) * ratio - event.y)
            self._update_adjustments()
            self.scroll_x = max(0, min(self.scroll_x, int(self.hadj.get_upper() - self.hadj.get_page_size())))
            self.scroll_y = max(0, min(self.scroll_y, int(self.vadj.get_upper() - self.vadj.get_page_size())))
            self.hadj.set_value(self.scroll_x)
            self.vadj.set_value(self.scroll_y)
            if self.on_zoom_changed:
                self.on_zoom_changed(self.zoom)
            self.drawing_area.queue_draw()
            return True

        step = 40
        if event.direction == Gdk.ScrollDirection.SMOOTH:
            dx, dy = _scroll_deltas(event)
            if event.state & Gdk.ModifierType.SHIFT_MASK:
                self.hadj.set_value(self.hadj.get_value() + dy * step)
            else:
                self.vadj.set_value(self.vadj.get_value() + dy * step)
                if dx:
                    self.hadj.set_value(self.hadj.get_value() + dx * step)
            return True
        if event.direction == Gdk.ScrollDirection.UP:
            self.vadj.set_value(self.vadj.get_value() - step)
        elif event.direction == Gdk.ScrollDirection.DOWN:
            self.vadj.set_value(self.vadj.get_value() + step)
        elif event.direction == Gdk.ScrollDirection.LEFT:
            self.hadj.set_value(self.hadj.get_value() - step)
        elif event.direction == Gdk.ScrollDirection.RIGHT:
            self.hadj.set_value(self.hadj.get_value() + step)
        return True

    def _on_key_press(self, widget, event):
        extend = bool(event.state & Gdk.ModifierType.SHIFT_MASK)
        ctrl = bool(event.state & Gdk.ModifierType.CONTROL_MASK)
        r, c = self.sel_cursor
        moved = True
        if event.keyval in (Gdk.KEY_Up, Gdk.KEY_KP_Up):
            r = max(1, r - 1)
        elif event.keyval in (Gdk.KEY_Down, Gdk.KEY_KP_Down):
            r = min(self.max_row, r + 1)
        elif event.keyval in (Gdk.KEY_Left, Gdk.KEY_KP_Left):
            c = max(1, c - 1)
        elif event.keyval in (Gdk.KEY_Right, Gdk.KEY_KP_Right):
            c = min(self.max_col, c + 1)
        elif event.keyval == Gdk.KEY_Home:
            c = 1
            if ctrl:
                r = 1
        elif event.keyval == Gdk.KEY_End:
            c = self.max_col
            if ctrl:
                r = self.max_row
        elif event.keyval in (Gdk.KEY_Page_Up, Gdk.KEY_Page_Down):
            alloc = self.drawing_area.get_allocation()
            page = max(1, (alloc.height - self.header_row_height) // DEFAULT_ROW_HEIGHT)
            if event.keyval == Gdk.KEY_Page_Up:
                r = max(1, r - page)
            else:
                r = min(self.max_row, r + page)
        else:
            moved = False

        if moved:
            self.sel_cursor = (r, c)
            if not extend:
                self.sel_anchor = (r, c)
            self._scroll_to_cell(r, c)
            self._notify_selection()
            self.drawing_area.queue_draw()
            return True
        return False

    def _scroll_to_cell(self, row, col):
        alloc = self.drawing_area.get_allocation()
        x_doc = self._col_x(col)
        y_doc = self._row_y(row)
        w = self.col_widths[col - 1]
        h = self.row_heights[row - 1]
        if x_doc < self.scroll_x + self.header_col_width:
            self.hadj.set_value(max(0, x_doc - self.header_col_width))
        elif x_doc + w > self.scroll_x + alloc.width:
            self.hadj.set_value(x_doc + w - alloc.width)
        if y_doc < self.scroll_y + self.header_row_height:
            self.vadj.set_value(max(0, y_doc - self.header_row_height))
        elif y_doc + h > self.scroll_y + alloc.height:
            self.vadj.set_value(y_doc + h - alloc.height)

    def _notify_selection(self):
        if self.on_selection_changed:
            self.on_selection_changed(self)

    def get_selection_text(self):
        t, l, b, r = self._sel_rect()
        lines = []
        for row in range(t, b + 1):
            fields = []
            for col in range(l, r + 1):
                if (row, col) in self.merged_covered:
                    ar, ac = self.merged_covered[(row, col)]
                    value = self.ws.cell(row=ar, column=ac).value
                else:
                    value = self.ws.cell(row=row, column=col).value
                fields.append(format_cell_value(value))
            lines.append("\t".join(fields))
        return "\n".join(lines)

    def selection_info(self):
        t, l, b, r = self._sel_rect()
        if t == b and l == r:
            return f"{get_column_letter(l)}{t}"
        return f"{get_column_letter(l)}{t}:{get_column_letter(r)}{b}"

    def selection_stats(self):
        t, l, b, r = self._sel_rect()
        count = 0
        num_count = 0
        total = 0.0
        vmin = None
        vmax = None
        for row in range(t, b + 1):
            for col in range(l, r + 1):
                if (row, col) in self.merged_covered:
                    continue
                v = self.ws.cell(row=row, column=col).value
                if v is None or v == "":
                    continue
                count += 1
                if isinstance(v, bool):
                    continue
                if isinstance(v, (int, float)):
                    num_count += 1
                    total += v
                    vmin = v if vmin is None or v < vmin else vmin
                    vmax = v if vmax is None or v > vmax else vmax
        return {
            "rows": b - t + 1,
            "cols": r - l + 1,
            "count": count,
            "num_count": num_count,
            "sum": total if num_count else None,
            "avg": total / num_count if num_count else None,
            "min": vmin,
            "max": vmax,
        }

    def anchor_value_text(self):
        t, l, _, _ = self._sel_rect()
        if (t, l) in self.merged_covered:
            ar, ac = self.merged_covered[(t, l)]
            v = self.ws.cell(row=ar, column=ac).value
        else:
            v = self.ws.cell(row=t, column=l).value
        return format_cell_value(v)

    def set_zoom(self, z):
        old = self.zoom
        self.zoom = max(MIN_ZOOM, min(MAX_ZOOM, z))
        if abs(self.zoom - old) < 1e-6:
            return
        self._update_adjustments()
        self.drawing_area.queue_draw()
        if self.on_zoom_changed:
            self.on_zoom_changed(self.zoom)


class QuickCellApp:
    def __init__(self):
        self.window = Gtk.Window(title="QuickCell")
        self.window.set_default_size(1100, 720)
        self.window.connect("destroy", Gtk.main_quit)
        self.window.connect("key-press-event", self._on_window_key)

        self.filepath = None
        self.wb = None
        self.sheet_views = []

        open_btn = Gtk.Button(label="📂 Open")
        open_btn.connect("clicked", lambda *_: self.open_dialog())
        copy_btn = Gtk.Button(label="📄 Copy")
        copy_btn.connect("clicked", lambda *_: self.copy_selection())
        zout_btn = Gtk.Button(label="🔍−")
        zout_btn.connect("clicked", lambda *_: self.zoom_delta(1 / 1.1))
        zreset_btn = Gtk.Button(label="1:1")
        zreset_btn.connect("clicked", lambda *_: self.zoom_set(1.0))
        zin_btn = Gtk.Button(label="🔍+")
        zin_btn.connect("clicked", lambda *_: self.zoom_delta(1.1))
        help_btn = Gtk.Button(label="❓ Help")
        help_btn.connect("clicked", lambda *_: self.show_help())

        hbox = Gtk.Box(spacing=6)
        hbox.pack_start(open_btn, False, False, 0)
        hbox.pack_start(copy_btn, False, False, 0)
        hbox.pack_start(Gtk.Separator(orientation=Gtk.Orientation.VERTICAL), False, False, 0)
        hbox.pack_start(zout_btn, False, False, 0)
        hbox.pack_start(zreset_btn, False, False, 0)
        hbox.pack_start(zin_btn, False, False, 0)
        hbox.pack_start(Gtk.Separator(orientation=Gtk.Orientation.VERTICAL), False, False, 0)
        hbox.pack_start(help_btn, False, False, 0)

        self.notebook = Gtk.Notebook()
        self.notebook.set_scrollable(True)
        self.notebook.connect("switch-page", self._on_switch_page)

        self.status_cell = Gtk.Label(label="")
        self.status_cell.set_xalign(0)
        self.status_cell.set_size_request(110, -1)
        self.status_value = Gtk.Label(label="")
        self.status_value.set_xalign(0)
        self.status_value.set_ellipsize(Pango.EllipsizeMode.END)
        self.status_zoom = Gtk.Label(label="100%")
        self.status_zoom.set_xalign(1)
        self.status_zoom.set_size_request(60, -1)

        status_box = Gtk.Box(spacing=12)
        status_box.set_margin_start(8)
        status_box.set_margin_end(8)
        status_box.set_margin_top(2)
        status_box.set_margin_bottom(2)
        status_box.pack_start(self.status_cell, False, False, 0)
        status_box.pack_start(self.status_value, True, True, 0)
        status_box.pack_start(self.status_zoom, False, False, 0)

        self.toast_label = Gtk.Label()
        css = Gtk.CssProvider()
        css.load_from_data(
            b".toast { background-color: rgba(0,0,0,0.8); color: white; "
            b"padding: 10px 20px; border-radius: 5px; }"
        )
        self.toast_label.get_style_context().add_provider(
            css, Gtk.STYLE_PROVIDER_PRIORITY_APPLICATION
        )
        self.toast_label.get_style_context().add_class("toast")
        self.toast_label.set_halign(Gtk.Align.CENTER)
        self.toast_label.set_valign(Gtk.Align.START)
        self.toast_label.set_margin_top(10)
        self.toast_label.set_no_show_all(True)
        self._toast_timer = None

        self.overlay = Gtk.Overlay()
        self.overlay.add_overlay(self.toast_label)

        vbox = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=4)
        vbox.set_margin_top(4)
        vbox.set_margin_bottom(4)
        vbox.set_margin_start(4)
        vbox.set_margin_end(4)
        vbox.pack_start(hbox, False, False, 0)
        vbox.pack_start(self.notebook, True, True, 0)
        vbox.pack_start(status_box, False, False, 0)

        self.overlay.add(vbox)
        self.window.add(self.overlay)
        self.window.show_all()

        self._show_empty_state()

    def _show_empty_state(self):
        placeholder = Gtk.Label(
            label="Open an .xlsx file (Ctrl+O, or pass a path on the command line)"
        )
        placeholder.set_vexpand(True)
        placeholder.set_hexpand(True)
        self.notebook.append_page(placeholder, Gtk.Label(label="(no file)"))
        self.notebook.show_all()

    def show_toast(self, message):
        if self._toast_timer:
            GLib.source_remove(self._toast_timer)
        self.toast_label.set_text(message)
        self.toast_label.show()
        self._toast_timer = GLib.timeout_add(2000, self._hide_toast)

    def _hide_toast(self):
        self._toast_timer = None
        self.toast_label.hide()
        return False

    def load_file(self, path):
        if not os.path.exists(path):
            self.show_toast(f"File not found: {os.path.basename(path)}")
            return False
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as e:
            self.show_toast(f"Cannot open: {e}")
            return False

        while self.notebook.get_n_pages() > 0:
            self.notebook.remove_page(-1)
        self.sheet_views = []

        self.wb = wb
        self.filepath = path
        for name in wb.sheetnames:
            ws = wb[name]
            view = SheetView(
                ws,
                on_selection_changed=self._on_selection_changed,
                on_zoom_changed=self._on_zoom_changed,
            )
            self.sheet_views.append(view)
            self.notebook.append_page(view, Gtk.Label(label=name))

        self.notebook.show_all()
        self.window.set_title(f"QuickCell — {os.path.basename(path)}")
        self.notebook.set_current_page(0)
        if self.sheet_views:
            v = self.sheet_views[0]
            GLib.idle_add(v.drawing_area.grab_focus)
            self._on_selection_changed(v)
            self._on_zoom_changed(v.zoom)
        return True

    def open_dialog(self):
        dialog = Gtk.FileChooserDialog(
            title="Open Excel file",
            parent=self.window,
            action=Gtk.FileChooserAction.OPEN,
        )
        dialog.add_button("Cancel", Gtk.ResponseType.CANCEL)
        dialog.add_button("Open", Gtk.ResponseType.OK)
        flt = Gtk.FileFilter()
        flt.set_name("Excel files")
        flt.add_pattern("*.xlsx")
        flt.add_pattern("*.xlsm")
        dialog.add_filter(flt)
        any_flt = Gtk.FileFilter()
        any_flt.set_name("All files")
        any_flt.add_pattern("*")
        dialog.add_filter(any_flt)
        resp = dialog.run()
        path = dialog.get_filename() if resp == Gtk.ResponseType.OK else None
        dialog.destroy()
        if path:
            self.load_file(path)

    def current_view(self):
        page = self.notebook.get_current_page()
        if 0 <= page < len(self.sheet_views):
            return self.sheet_views[page]
        return None

    def _on_switch_page(self, notebook, page, page_num):
        if 0 <= page_num < len(self.sheet_views):
            v = self.sheet_views[page_num]
            self._on_selection_changed(v)
            self._on_zoom_changed(v.zoom)
            GLib.idle_add(v.drawing_area.grab_focus)

    def _on_selection_changed(self, view):
        if view is not self.current_view():
            return
        self.status_cell.set_text(view.selection_info())
        stats = view.selection_stats()
        if stats["rows"] == 1 and stats["cols"] == 1:
            self.status_value.set_text(view.anchor_value_text())
            return
        parts = [f"{stats['rows']}R × {stats['cols']}C"]
        parts.append(f"Count: {stats['count']}")
        if stats["num_count"]:
            parts.append(f"Sum: {_fmt_stat(stats['sum'])}")
            parts.append(f"Avg: {_fmt_stat(stats['avg'])}")
            parts.append(f"Min: {_fmt_stat(stats['min'])}")
            parts.append(f"Max: {_fmt_stat(stats['max'])}")
        self.status_value.set_text("   ".join(parts))

    def _on_zoom_changed(self, zoom):
        self.status_zoom.set_text(f"{int(round(zoom * 100))}%")

    def copy_selection(self):
        view = self.current_view()
        if view is None:
            return
        text = view.get_selection_text()
        clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)
        clipboard.set_text(text, -1)
        self.show_toast(f"✓ Copied {view.selection_info()}")

    def zoom_delta(self, factor):
        view = self.current_view()
        if view is None:
            return
        view.set_zoom(view.zoom * factor)

    def zoom_set(self, z):
        view = self.current_view()
        if view is None:
            return
        view.set_zoom(z)

    def _on_window_key(self, widget, event):
        ctrl = bool(event.state & Gdk.ModifierType.CONTROL_MASK)
        if not ctrl:
            return False
        k = event.keyval
        if k == Gdk.KEY_c:
            self.copy_selection()
            return True
        if k == Gdk.KEY_o:
            self.open_dialog()
            return True
        if k in (Gdk.KEY_plus, Gdk.KEY_equal, Gdk.KEY_KP_Add):
            self.zoom_delta(1.1)
            return True
        if k in (Gdk.KEY_minus, Gdk.KEY_KP_Subtract):
            self.zoom_delta(1 / 1.1)
            return True
        if k == Gdk.KEY_0:
            self.zoom_set(1.0)
            return True
        if k == Gdk.KEY_Page_Down:
            self.notebook.next_page()
            return True
        if k == Gdk.KEY_Page_Up:
            self.notebook.prev_page()
            return True
        return False

    def show_help(self):
        dialog = Gtk.Dialog(title="Help", parent=self.window, flags=0)
        dialog.add_button("Close", Gtk.ResponseType.CLOSE)
        text = Gtk.Label()
        text.set_markup(
            f"""<b>QuickCell v{VERSION}</b>

A minimal read-only viewer for .xlsx files.

<b>Mouse:</b>
• Click / drag: select a cell or range
• Shift+click: extend selection
• Click column/row header: select whole column/row
• Click top-left corner: select all

<b>Keyboard:</b>
• Arrow keys / Home / End / PgUp / PgDn: navigate
• Shift+arrows: extend selection
• Ctrl+Home / Ctrl+End: jump to A1 / last cell
• Ctrl+C: copy selection (tab-separated)
• Ctrl+O: open file
• Ctrl+scroll / Ctrl++ / Ctrl+− / Ctrl+0: zoom
• Ctrl+PgUp / Ctrl+PgDn: switch sheet

<b>Notes:</b>
• Files are opened read-only; nothing is ever written back.
• Formula results come from the cached value stored in the file
  (what Excel/LibreOffice wrote last time it saved)."""
        )
        text.set_margin_start(15)
        text.set_margin_end(15)
        text.set_margin_top(15)
        text.set_margin_bottom(15)
        dialog.get_content_area().add(text)
        dialog.show_all()
        dialog.run()
        dialog.destroy()

    def _load_initial_file(self, path):
        self.load_file(path)
        return GLib.SOURCE_REMOVE


def main():
    app = QuickCellApp()
    if len(sys.argv) > 1:
        GLib.idle_add(app._load_initial_file, sys.argv[1])
    Gtk.main()


if __name__ == "__main__":
    main()
