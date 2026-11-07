"""Gera controle_prazos.fods (XML do LibreOffice) sem bibliotecas externas.
Depois converta para .ods:  soffice --headless --convert-to ods controle_prazos.fods
"""
from xml.sax.saxutils import escape

PESSOAS = [
    ("Ana Souza",      "Analista",   "Atraso na entrega do relatório", "2026-10-28"),
    ("Bruno Lima",     "Técnico",    "Falta sem justificativa",        "2026-11-02"),
    ("Carla Dias",     "Supervisor", "Descumprimento de procedimento", "2026-11-03"),
    ("Daniel Rocha",   "Analista",   "Uso indevido de equipamento",    "2026-10-20"),
    ("Elisa Martins",  "Técnico",    "Atraso na entrega do relatório", "2026-11-05"),
    ("Felipe Costa",   "Estagiário", "Falta sem justificativa",        "2026-10-30"),
    ("Gabriela Nunes", "Gerente",    "Descumprimento de procedimento", "2026-11-04"),
    ("Hugo Ferreira",  "Analista",   "Uso indevido de equipamento",    "2026-11-01"),
]

def txt(v): return f'<table:table-cell office:value-type="string"><text:p>{escape(v)}</text:p></table:table-cell>'
def dt(v):  return f'<table:table-cell table:style-name="ce-data" office:value-type="date" office:date-value="{v}"><text:p>{v}</text:p></table:table-cell>'
def row(cells): return "<table:table-row>" + "".join(cells) + "</table:table-row>"

pessoas = [row(txt(h) for h in ["Nome","Cargo","Ocorrência","Data"])]
pessoas += [row([txt(n),txt(c),txt(o),dt(d)]) for n,c,o,d in PESSOAS]
prazos = [row(txt(h) for h in ["Nome","Cargo","Ocorrência","Data","Prazo Limite","Situação"])]

doc = f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
<office:automatic-styles>
 <number:date-style style:name="N-data"><number:day number:style="long"/><number:text>/</number:text><number:month number:style="long"/><number:text>/</number:text><number:year number:style="long"/></number:date-style>
 <style:style style:name="ce-data" style:family="table-cell" style:data-style-name="N-data"/>
</office:automatic-styles>
<office:body><office:spreadsheet>
<table:table table:name="Pessoas">{"".join(pessoas)}</table:table>
<table:table table:name="Prazos">{"".join(prazos)}</table:table>
</office:spreadsheet></office:body></office:document>'''
open("controle_prazos.fods","w",encoding="utf-8").write(doc)
print("ok")
