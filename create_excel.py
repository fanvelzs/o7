from typing import List

def create_excel(path: str):
    header = """<?xml version='1.0'?>\n<?mso-application progid='Excel.Sheet'?>\n<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet'\n xmlns:o='urn:schemas-microsoft-com:office:office'\n xmlns:x='urn:schemas-microsoft-com:office:excel'\n xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'\n xmlns:html='http://www.w3.org/TR/REC-html40'>\n"""
    footer = "</Workbook>\n"

    pagos_sheet = [
        "  <Worksheet ss:Name='Pagos'>",
        "    <Table>",
        "      <Row>",
        "        <Cell><Data ss:Type='String'>Numero de Factura</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Cliente/Empresa</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Monto Total</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Abono</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Fecha de Abono</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Saldo</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Estado de Pago</Data></Cell>",
        "      </Row>"
    ]
    # add empty rows with formulas for saldo and estado de pago
    for i in range(2, 52):
        saldo_formula = f"=C{i}-SUMIFS($D$2:$D$1000,$A$2:$A$1000,A{i})"
        estado_formula = f"=IF(F{i}<=0,\"Pagado\",\"Pendiente\")"
        pagos_sheet.append("      <Row>")
        pagos_sheet.append("        <Cell><Data ss:Type='String'></Data></Cell>")
        pagos_sheet.append("        <Cell><Data ss:Type='String'></Data></Cell>")
        pagos_sheet.append("        <Cell><Data ss:Type='Number'></Data></Cell>")
        pagos_sheet.append("        <Cell><Data ss:Type='Number'></Data></Cell>")
        pagos_sheet.append("        <Cell><Data ss:Type='String'></Data></Cell>")
        pagos_sheet.append(f"        <Cell ss:Formula='{saldo_formula}'><Data ss:Type='Number'>0</Data></Cell>")
        pagos_sheet.append(f"        <Cell ss:Formula='{estado_formula}'><Data ss:Type='String'></Data></Cell>")
        pagos_sheet.append("      </Row>")
    pagos_sheet.append("    </Table>")
    pagos_sheet.append("  </Worksheet>")

    resumen_sheet = [
        "  <Worksheet ss:Name='HistorialPorCliente'>",
        "    <Table>",
        "      <Row>",
        "        <Cell><Data ss:Type='String'>Cliente/Empresa</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Numero de Factura</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Monto Total</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Total Abono</Data></Cell>",
        "        <Cell><Data ss:Type='String'>Saldo Restante</Data></Cell>",
        "      </Row>"
    ]
    for i in range(2, 52):
        total_formula = f"=SUMIFS(Pagos!$C$2:$C$1000,Pagos!$A$2:$A$1000,$B{i},Pagos!$B$2:$B$1000,$A{i})"
        abono_formula = f"=SUMIFS(Pagos!$D$2:$D$1000,Pagos!$A$2:$A$1000,$B{i},Pagos!$B$2:$B$1000,$A{i})"
        saldo_formula = f"=C{i}-D{i}"
        resumen_sheet.append("      <Row>")
        resumen_sheet.append("        <Cell><Data ss:Type='String'></Data></Cell>")
        resumen_sheet.append("        <Cell><Data ss:Type='String'></Data></Cell>")
        resumen_sheet.append(f"        <Cell ss:Formula='{total_formula}'><Data ss:Type='Number'>0</Data></Cell>")
        resumen_sheet.append(f"        <Cell ss:Formula='{abono_formula}'><Data ss:Type='Number'>0</Data></Cell>")
        resumen_sheet.append(f"        <Cell ss:Formula='{saldo_formula}'><Data ss:Type='Number'>0</Data></Cell>")
        resumen_sheet.append("      </Row>")
    resumen_sheet.append("    </Table>")
    resumen_sheet.append("  </Worksheet>")

    content = header + "\n".join(pagos_sheet) + "\n" + "\n".join(resumen_sheet) + "\n" + footer
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == '__main__':
    create_excel('control_pagos.xls')
