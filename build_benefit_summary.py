from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Predefined styles (can be moved to config if needed)
HEADER_FILL = PatternFill(start_color="266B8E", end_color="266B8E", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Calibri")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Common benefit categories in row order
BENEFIT_ROWS = [
    "Deductible (Individual)",
    "Deductible (Family)",
    "Out-of-Pocket Max (Individual)",
    "Out-of-Pocket Max (Family)",
    "PCP Visit",
    "Specialist Visit",
    "Urgent Care",
    "Emergency Room",
    "Generic Rx",
    "Brand Rx",
    "Specialty Rx",
    "Hospitalization",
    "Mental Health",
    "Telehealth"
]

def build_benefit_summary(ws, plans, start_row=2, start_col=2):
    """
    ws: openpyxl worksheet object
    plans: list of dicts, each with keys: 'name', 'benefits' (dict of category: value)
    """
    if not plans:
        return

    # Write headers (plan names)
    for i, plan in enumerate(plans):
        col = start_col + i
        cell = ws.cell(row=start_row, column=col, value=plan['name'])
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

    # Write benefit categories + plan values
    for j, category in enumerate(BENEFIT_ROWS):
        label_cell = ws.cell(row=start_row + 1 + j, column=start_col - 1, value=category)
        label_cell.font = Font(bold=True, name="Calibri")
        label_cell.alignment = Alignment(horizontal="left")
        label_cell.border = THIN_BORDER

        for i, plan in enumerate(plans):
            col = start_col + i
            value = plan['benefits'].get(category, "N/A")
            cell = ws.cell(row=start_row + 1 + j, column=col, value=value)
            cell.alignment = CENTER_ALIGN
            cell.border = THIN_BORDER

    # Auto-width columns
    for i in range(start_col - 1, start_col + len(plans)):
        ws.column_dimensions[get_column_letter(i + 1)].width = 22
