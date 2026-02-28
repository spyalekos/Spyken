with open('main.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find the problematic ft.Text block (lines 247-251, 0-indexed 246-250)
# Replace lines 246-250 with a fixed version
fix = (
    '                    ft.Text(\n'
    '                        "\u03a4\u03bf Spyken \u03bc\u03b5\u03c4\u03b1\u03c4\u03c1\u03ad\u03c0\u03b5\u03b9 \u03b1\u03c1\u03c7\u03b5\u03af\u03b1 .docx \u03ba\u03b1\u03b9 .pdf \u03c3\u03b5 \u03b1\u03c1\u03c7\u03b5\u03af\u03b1 \u03ae\u03c7\u03bf\u03c5 .mp3, "'
    '"\u03b5\u03bd\u03b1\u03bb\u03bb\u03ac\u03c3\u03c3\u03bf\u03bd\u03c4\u03b1\u03c2 \u03b1\u03bd\u03b4\u03c1\u03b9\u03ba\u03ae \u03ba\u03b1\u03b9 \u03b3\u03c5\u03bd\u03b1\u03b9\u03ba\u03b5\u03af\u03b1 \u03c6\u03c9\u03bd\u03ae \u03b3\u03b9\u03b1 \u03ba\u03ac\u03b8\u03b5 \u03c0\u03b1\u03c1\u03ac\u03b3\u03c1\u03b1\u03c6\u03bf.",\n'
    '                        size=14, color=ft.Colors.GREY_300\n'
    '                    ),\n'
)

new_lines = lines[:246] + [fix] + lines[251:]

with open('main.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Done, total lines:", len(new_lines))
