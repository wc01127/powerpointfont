import sys
from powerpointfont import set_uniform_font, analyze_fonts

if len(sys.argv) != 3:
    print("Usage: python process_pptx.py <input_pptx> <target_font>")
    sys.exit(1)

input_pptx = sys.argv[1]
target_font = sys.argv[2]
output_pptx = input_pptx.rsplit('.', 1)[0] + '_updated.pptx'

# Analyze fonts
analyze_fonts(input_pptx)

# Uniformize fonts
changes_made, fonts_changed = set_uniform_font(
    input_pptx, output_pptx, target_font)

# Prepare output
output = f"File processed: {input_pptx}\n"
output += f"Target font: {target_font}\n\n"

output += "Font Uniformization:\n"
if changes_made:
    output += f"Fonts updated to '{target_font}'. Saved as '{output_pptx}'\n"
    output += "Font changes summary:\n"
    for font, count in fonts_changed.items():
        output += f"  - {font} to {target_font}: {count} occurrence{'s' if count > 1 else ''}\n"
    if "Default font" in fonts_changed:
        output += "  Note: Some 'Default font' text was changed to the target font.\n"
else:
    output += f"No font changes were necessary. All fonts are already '{target_font}'.\n"

print(output)
