"""Rebuild Productos Vendidos + Reservado formulas with correct column refs per sheet."""
import gspread

gc = gspread.service_account(filename=r'C:\Users\tadeu\maleu-service-account.json')
sh = gc.open_by_key('1ILXCc9ddbC_gJPNoUADBiSMXAWLM9v73ov2_xXb8YsY')
prod = sh.worksheet('Productos')

def col_letter(n):
    s = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s

# abbr -> col number per sheet (1-based). None = not in sheet.
HOME = {
    'PMu': 37, 'PMa': 38, 'PJyQ': 39, 'PCC': 40, 'PJyM': 41,
    'PPM': 23, 'PPJyQ': 24, 'PPCyQ': 25,
    'SCo': 26, 'SJyQ': 27, 'SCa': 28,
    'ECaC': 29, 'EJyQ': 30, 'ECyQ': 31, 'EV': 32,
    'TG': 33, 'TLC': 34, 'TC': 35, 'F': 36,
}
PILAR = {
    'PMu': 41, 'PMa': 42, 'PJyQ': 43, 'PCC': 44, 'PJyM': 45,
    'PPM': 23, 'PPJyQ': 24, 'PPCyQ': 25,
    'SQB': 26, 'SL': 27, 'SCo': 28, 'SPyP': 29, 'SJyQ': 30, 'SE': 31, 'SCa': 32,
    'ECaC': 33, 'EJyQ': 34, 'ECyQ': 35, 'EV': 36,
    'TG': 37, 'TLC': 38, 'TC': 39, 'F': 40,
}
CLUBES = {
    'PMu': 24, 'PMa': 25, 'PJyQ': 26, 'PCC': 27, 'PJyM': 28,
    'PPM': 29, 'PPJyQ': 30, 'PPCyQ': 31,
}

# Read abbreviaturas from Productos
abbrs = prod.get('C2:C24')
for i, r in enumerate(abbrs):
    row_n = i + 2
    abbr = r[0] if r else ''
    if not abbr:
        continue

    # Build Vendidos formula (filter: Deposito + Entregado + Semana/Año actual)
    v_parts = []
    r_parts = []
    if abbr in HOME:
        L = col_letter(HOME[abbr])
        v_parts.append(f'SUMPRODUCT((Home!$I$2:$I$9993="Deposito")*(Home!$K$2:$K$9993="Entregado")*(Home!$AZ$2:$AZ$9993=ISOWEEKNUM(TODAY()))*(Home!$BA$2:$BA$9993=YEAR(TODAY()))*(Home!{L}$2:{L}$9993))')
        r_parts.append(f'SUMPRODUCT((Home!$I$2:$I$9993="Deposito")*(Home!$K$2:$K$9993="Reservado")*(Home!{L}$2:{L}$9993))')
    if abbr in PILAR:
        L = col_letter(PILAR[abbr])
        v_parts.append(f'SUMPRODUCT((Pilar!$I$2:$I$9994="Deposito")*(Pilar!$K$2:$K$9994="Entregado")*(Pilar!$BC$2:$BC$9994=ISOWEEKNUM(TODAY()))*(Pilar!$BD$2:$BD$9994=YEAR(TODAY()))*(Pilar!{L}$2:{L}$9994))')
        r_parts.append(f'SUMPRODUCT((Pilar!$I$2:$I$9994="Deposito")*(Pilar!$K$2:$K$9994="Reservado")*(Pilar!{L}$2:{L}$9994))')
    if abbr in CLUBES:
        L = col_letter(CLUBES[abbr])
        v_parts.append(f'SUMPRODUCT((Clubes!$L$2:$L$9996="Deposito")*(Clubes!$N$2:$N$9996="Entregado")*(Clubes!$F$2:$F$9996=ISOWEEKNUM(TODAY()))*(Clubes!$G$2:$G$9996=YEAR(TODAY()))*(Clubes!{L}$2:{L}$9996))')
        r_parts.append(f'SUMPRODUCT((Clubes!$L$2:$L$9996="Deposito")*(Clubes!$N$2:$N$9996="Reservado")*(Clubes!{L}$2:{L}$9996))')

    vendidos = '=' + '+'.join(v_parts) if v_parts else '=0'
    reservado = '=' + '+'.join(r_parts) if r_parts else '=0'

    prod.update(range_name=f'E{row_n}', values=[[vendidos]], value_input_option='USER_ENTERED')
    prod.update(range_name=f'G{row_n}', values=[[reservado]], value_input_option='USER_ENTERED')
    print(f'row {row_n} ({abbr}): Vendidos + Reservado actualizados')

print('\nListo.')
