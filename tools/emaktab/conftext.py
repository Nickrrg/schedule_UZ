# Выгрузка Confluence (.doc = MIME multipart с HTML) → текст с таблицами через « | ».
import sys, io, email, re, html
from html.parser import HTMLParser
msg = email.message_from_bytes(open(sys.argv[1], 'rb').read())
parts = []
for part in msg.walk():
    if part.get_content_type() == 'text/html':
        raw = part.get_payload(decode=True)
        cs = part.get_content_charset() or 'utf-8'
        parts.append(raw.decode(cs, 'ignore'))
src = '\n'.join(parts)

class P(HTMLParser):
    def __init__(self):
        super().__init__(); self.out = []; self.skip = 0
    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style'): self.skip += 1
        if tag in ('p', 'div', 'br', 'tr', 'li', 'h1', 'h2', 'h3', 'h4', 'table'): self.out.append('\n')
        if tag in ('h1', 'h2', 'h3', 'h4'): self.out.append('#' * int(tag[1]) + ' ')
        if tag == 'li': self.out.append('- ')
        if tag in ('td', 'th'): self.out.append(' | ')
        if tag == 'img':
            a = dict(attrs); self.out.append('[img ' + (a.get('alt') or a.get('src', ''))[:60] + ']')
    def handle_endtag(self, tag):
        if tag in ('script', 'style'): self.skip -= 1
    def handle_data(self, d):
        if not self.skip: self.out.append(d)
p = P(); p.feed(src)
text = ''.join(p.out)
text = re.sub(r'[ \t\xa0]+', ' ', text)
text = re.sub(r'\n\s*\n+', '\n\n', text)
io.open(sys.argv[2], 'w', encoding='utf-8').write(text.strip())
print(len(text))
