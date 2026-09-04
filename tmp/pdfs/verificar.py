from PIL import Image, ImageOps, ImageDraw
from pypdf import PdfReader
from pathlib import Path
files=sorted(Path('tmp/pdfs').glob('pagina-*.png'))
thumbs=[]
for f in files:
 im=Image.open(f).convert('RGB'); im.thumbnail((450,650)); thumbs.append(im)
board=Image.new('RGB',(450*len(thumbs),650),'#bbbbbb')
for i,im in enumerate(thumbs): board.paste(im,(450*i,0))
board.save('tmp/pdfs/revision.png')
r=PdfReader('output/pdf/instructivo_portal_del_jugador.pdf')
print('Pages:',len(r.pages))
for i,page in enumerate(r.pages): print(i+1,len(page.extract_text()))
