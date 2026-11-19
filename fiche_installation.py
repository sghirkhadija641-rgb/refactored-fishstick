"""Fiche A4 « Installer l'application sur votre téléphone » : ajout des deux formulaires Ré-Use
sur l'écran d'accueil (iPhone et Android). PDF avec liens cliquables + aperçu PNG."""
import pymupdf
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white

SORTIE = 'fiche-installation-reuse'
LIENS = [("Saisie d'une réclamation", "https://tinyurl.com/reuse-reclamations", "Réclamations Ré-Use"),
         ("Traitement de réclamation", "https://tinyurl.com/reuse-traitement", "Traitement Ré-Use")]

pdfmetrics.registerFont(TTFont('B', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
pdfmetrics.registerFont(TTFont('R', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
teal = HexColor('#0f6e86'); ink = HexColor('#14262c'); muted = HexColor('#5b7078'); line = HexColor('#d6e0e3'); soft = HexColor('#f3f6f7')
W, H = A4
c = canvas.Canvas(SORTIE + '.pdf', pagesize=A4); c.setTitle("Installer l'application Ré-Use")

c.setFillColor(teal); c.rect(0, H-120, W, 120, stroke=0, fill=1)
c.setFillColor(white); c.setFont('B', 26); c.drawCentredString(W/2, H-62, "Installer l'application Ré-Use")
c.setFont('R', 13); c.setFillColor(HexColor('#dff1f5')); c.drawCentredString(W/2, H-90, "Une icône sur l'écran d'accueil du téléphone, en 30 secondes")

# Les deux liens
y = H - 150
c.setFillColor(ink); c.setFont('B', 13); c.drawString(40, y, "1. Ouvrez le lien sur votre téléphone (connecté à votre compte Veolia)")
y -= 14
for i, (titre, url, nom) in enumerate(LIENS):
    bx = 40 + i*(W-80)/2; bw = (W-80)/2 - 10; by = y - 62
    c.setFillColor(soft); c.setStrokeColor(line); c.roundRect(bx, by, bw, 54, 8, stroke=1, fill=1)
    c.setFillColor(teal); c.roundRect(bx, by, 5, 54, 2, stroke=0, fill=1)
    c.setFillColor(ink); c.setFont('B', 11.5); c.drawString(bx+16, by+34, titre)
    show = url.replace('https://', ''); c.setFillColor(teal); c.setFont('B', 11)
    tw = c.stringWidth(show, 'B', 11); c.drawString(bx+16, by+14, show)
    c.setStrokeColor(teal); c.setLineWidth(0.7); c.line(bx+16, by+11, bx+16+tw, by+11)
    c.linkURL(url, (bx, by, bx+bw, by+54), relative=0)
y -= 92

c.setFillColor(ink); c.setFont('B', 13); c.drawString(40, y, "2. Ajoutez-le à l'écran d'accueil")
y -= 16

def bloc(x, y, w, titre, etapes):
    h = 46 + len(etapes)*44
    c.setFillColor(white); c.setStrokeColor(line); c.roundRect(x, y-h, w, h, 10, stroke=1, fill=1)
    c.setFillColor(teal); c.roundRect(x, y-34, w, 34, 10, stroke=0, fill=1); c.rect(x, y-34, w, 12, stroke=0, fill=1)
    c.setFillColor(white); c.setFont('B', 14); c.drawString(x+16, y-23, titre)
    yy = y - 58
    for n, (fort, detail) in enumerate(etapes, 1):
        c.setFillColor(teal); c.circle(x+26, yy+4, 10, stroke=0, fill=1)
        c.setFillColor(white); c.setFont('B', 11); c.drawCentredString(x+26, yy, str(n))
        c.setFillColor(ink); c.setFont('B', 11); c.drawString(x+46, yy+3, fort)
        c.setFillColor(muted); c.setFont('R', 9.5); c.drawString(x+46, yy-11, detail)
        yy -= 44
    return h

colw = (W-80)/2 - 8
h1 = bloc(40, y, colw, "iPhone (Safari)", [
    ("Ouvrez le lien dans Safari", "Pas dans une autre application."),
    ("Touchez « Partager »", "Le carré avec une flèche vers le haut."),
    ("« Sur l'écran d'accueil »", "Faites défiler la liste si besoin."),
    ("Nommez-le puis « Ajouter »", "Ex. « Réclamations Ré-Use »."),
])
bloc(40 + colw + 16, y, colw, "Android (Chrome)", [
    ("Ouvrez le lien dans Chrome", "Le navigateur habituel d'Android."),
    ("Touchez « ⋮ »", "Les trois points en haut à droite."),
    ("« Ajouter à l'écran d'accueil »", "Ou « Installer l'application »."),
    ("Nommez-le puis « Ajouter »", "Ex. « Traitement Ré-Use »."),
])
y -= h1 + 34

c.setFillColor(ink); c.setFont('B', 13); c.drawString(40, y, "3. C'est prêt")
c.setFont('R', 11); c.setFillColor(muted)
for k, l in enumerate(["Une icône apparaît sur l'écran d'accueil : un appui ouvre directement le formulaire.",
                       "Faites l'opération deux fois pour avoir les deux icônes : Saisie et Traitement.",
                       "Restez connecté à votre compte Veolia sur le téléphone pour pouvoir envoyer."]):
    c.drawString(40, y-20-k*17, "•  " + l)

c.setFont('R', 9); c.setFillColor(muted)
c.drawCentredString(W/2, 22, "Les réclamations et traitements sont transmis uniquement au responsable Ré-Use.")
c.showPage(); c.save()
pymupdf.open(SORTIE + '.pdf')[0].get_pixmap(dpi=110).save(SORTIE + '.png')
print('fiche installation OK')
