import tkinter as tk
from tkinter import ttk


PALETTE = {
    # 이름은 기존 호출부 호환을 위해 유지하지만, 실제 색상은 초록·흰색 계열로만 구성합니다.
    'window': '#F1F7F2',
    'surface': '#FFFFFF',
    'blue': '#DDEEE1',
    'blue_active': '#C8E1CE',
    'mint': '#E8F4EA',
    'mint_active': '#D4E9D8',
    'lavender': '#F4F8F4',
    'lavender_active': '#E4EFE6',
    'peach': '#EDF6EF',
    'rose': '#F8FBF8',
    'rose_active': '#E7F1E9',
    'entry': '#FFFFFF',
    'entry_disabled': '#E2E6E3',
    'text': '#263A2C',
    'muted': '#65776A',
    'border': '#B8D0BE',
    'primary': '#2F6842',
    'danger': '#496653',
    'success': '#2F6842',
    'strong': '#4F9661',
    'strong_active': '#3F7F50',
}


def apply_pastel_theme(root, font_size=10):
    """Windows에서도 색상이 일관되도록 clam 기반 파스텔 테마를 적용합니다."""
    root.configure(background=PALETTE['window'])
    root.option_add('*Font', ('Malgun Gothic', int(font_size)))
    root.option_add('*Listbox.background', PALETTE['entry'])
    root.option_add('*Listbox.foreground', PALETTE['text'])
    root.option_add('*Listbox.selectBackground', PALETTE['blue_active'])
    root.option_add('*Listbox.selectForeground', PALETTE['text'])

    style = ttk.Style(root)
    try:
        style.theme_use('clam')
    except tk.TclError:
        pass

    base_font = ('Malgun Gothic', int(font_size))
    style.configure('.', font=base_font, background=PALETTE['surface'], foreground=PALETTE['text'])
    style.configure('Window.TFrame', background=PALETTE['window'])
    style.configure('Card.TFrame', background=PALETTE['surface'])
    style.configure('TFrame', background=PALETTE['surface'])
    style.configure('TLabel', background=PALETTE['surface'], foreground=PALETTE['text'], padding=(0, 2))
    style.configure('Muted.TLabel', background=PALETTE['surface'], foreground=PALETTE['muted'])
    style.configure('TCheckbutton', background=PALETTE['surface'], foreground=PALETTE['text'], padding=(2, 3), indicatorbackground=PALETTE['entry'])
    style.map('TCheckbutton', background=[('active', PALETTE['mint'])], indicatorbackground=[('selected', PALETTE['strong'])])
    style.configure('TRadiobutton', background=PALETTE['surface'], foreground=PALETTE['text'])
    style.configure('TLabelframe', background=PALETTE['surface'], bordercolor=PALETTE['border'], borderwidth=1, relief='solid')
    style.configure('TLabelframe.Label', background=PALETTE['mint'], foreground=PALETTE['primary'], font=('Malgun Gothic', int(font_size), 'bold'), padding=(8, 4))
    style.configure('Card.TLabelframe', background=PALETTE['surface'], bordercolor=PALETTE['border'], borderwidth=1, relief='solid')
    style.configure('Card.TLabelframe.Label', background=PALETTE['mint'], foreground=PALETTE['primary'], font=('Malgun Gothic', int(font_size), 'bold'), padding=(8, 4))

    tones = {
        'Blue': PALETTE['blue'],
        'Mint': PALETTE['mint'],
        'Lavender': PALETTE['lavender'],
        'Peach': PALETTE['peach'],
    }
    for name, color in tones.items():
        style.configure(f'{name}.TLabelframe', background=PALETTE['surface'], bordercolor=PALETTE['border'], borderwidth=1, relief='solid')
        style.configure(f'{name}.TLabelframe.Label', background=color, foreground=PALETTE['primary'], font=('Malgun Gothic', int(font_size), 'bold'), padding=(8, 4))

    style.configure('TEntry', fieldbackground=PALETTE['entry'], foreground=PALETTE['text'], bordercolor=PALETTE['border'], lightcolor=PALETTE['border'], darkcolor=PALETTE['border'], padding=(7, 5))
    style.map(
        'TEntry',
        fieldbackground=[('disabled', PALETTE['entry_disabled']), ('readonly', PALETTE['entry_disabled'])],
        foreground=[('disabled', PALETTE['muted']), ('readonly', PALETTE['muted'])],
    )
    style.configure('TCombobox', fieldbackground=PALETTE['entry'], background=PALETTE['mint'], foreground=PALETTE['text'], arrowcolor=PALETTE['primary'], padding=(6, 4))
    style.map(
        'TCombobox',
        fieldbackground=[('disabled', PALETTE['entry_disabled']), ('readonly', PALETTE['entry'])],
        foreground=[('disabled', PALETTE['muted'])],
        selectbackground=[('disabled', PALETTE['entry_disabled']), ('readonly', PALETTE['entry'])],
        selectforeground=[('disabled', PALETTE['muted']), ('readonly', PALETTE['text'])],
    )
    style.configure('Vertical.TScrollbar', background=PALETTE['blue'], troughcolor=PALETTE['window'], bordercolor=PALETTE['window'], arrowcolor=PALETTE['primary'])
    style.configure('Horizontal.TScrollbar', background=PALETTE['blue'], troughcolor=PALETTE['window'], bordercolor=PALETTE['window'], arrowcolor=PALETTE['primary'])
    style.configure('TPanedwindow', background=PALETTE['window'], sashrelief='flat', sashwidth=6)
    style.configure('TSeparator', background=PALETTE['border'])
    return style


def apply_button_palette(kwargs):
    text = str(kwargs.get('text', ''))
    requested_fg = str(kwargs.get('fg', kwargs.get('foreground', ''))).lower()
    destructive = any(word in text for word in ('초기화', '삭제', '종료', '선정기준')) or requested_fg in ('#b00000', '#c00000', 'red')
    strong = any(word in text for word in ('첨부 파일 반영하기', '사용하기'))
    primary = any(word in text for word in ('저장', '반영', '사용하기', '확인', '선택'))
    if strong:
        background, active = PALETTE['strong'], PALETTE['strong_active']
        kwargs['fg'] = '#FFFFFF'
        if 'foreground' in kwargs:
            kwargs['foreground'] = '#FFFFFF'
    elif destructive:
        background, active = PALETTE['rose'], PALETTE['rose_active']
        kwargs['fg'] = PALETTE['danger']
        if 'foreground' in kwargs:
            kwargs['foreground'] = PALETTE['danger']
    elif primary:
        background, active = PALETTE['blue'], PALETTE['blue_active']
        kwargs.setdefault('fg', PALETTE['primary'])
    else:
        background, active = PALETTE['lavender'], PALETTE['lavender_active']
        kwargs.setdefault('fg', PALETTE['text'])
    kwargs.setdefault('background', background)
    kwargs.setdefault('activebackground', active)
    kwargs.setdefault('activeforeground', kwargs.get('fg', PALETTE['text']))
    kwargs.setdefault('highlightbackground', PALETTE['border'])
    kwargs.setdefault('highlightcolor', PALETTE['primary'])


def style_plain_widgets(root):
    """ttk 테마가 적용되지 않는 Canvas/Text/Listbox에도 같은 색을 입힙니다."""
    for widget in root.winfo_children():
        try:
            if isinstance(widget, tk.Canvas):
                widget.configure(background=PALETTE['window'])
            elif isinstance(widget, tk.Text):
                widget.configure(background=PALETTE['entry'], foreground=PALETTE['text'], insertbackground=PALETTE['text'], selectbackground=PALETTE['blue_active'], relief='solid', bd=1, padx=16, pady=14, spacing1=2, spacing3=2)
            elif isinstance(widget, tk.Listbox):
                widget.configure(background=PALETTE['entry'], foreground=PALETTE['text'], selectbackground=PALETTE['blue_active'], selectforeground=PALETTE['text'], relief='flat', bd=0, highlightthickness=1, highlightbackground=PALETTE['border'])
        except tk.TclError:
            pass
        style_plain_widgets(widget)
