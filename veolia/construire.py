"""Assemble Code.gs (un seul fichier à coller) : code du script + formulaire HTML intégré."""
import json, pathlib
d = pathlib.Path(__file__).parent
code = (d / 'code-modele.gs').read_text(encoding='utf8')
html = (d / 'formulaire.html').read_text(encoding='utf8')
(d / 'Code.gs').write_text(code.rstrip() + '\n\n// Formulaire affiché aux agents (généré depuis veolia/formulaire.html).\nconst FORM_HTML = ' + json.dumps(html, ensure_ascii=False) + ';\n', encoding='utf8')
print('Code.gs :', len((d / 'Code.gs').read_text(encoding='utf8')), 'caractères')
