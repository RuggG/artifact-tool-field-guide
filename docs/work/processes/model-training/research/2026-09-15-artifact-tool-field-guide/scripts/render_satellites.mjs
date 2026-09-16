// node render_satellites.mjs /absolute/path/to/marked/lib/marked.esm.js
// Render saved Markdown at build time, retaining the exact source alongside it.
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
const {marked} = await import(pathToFileURL(process.argv[2]).href);
const home = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const file = path.join(home, 'data/satellite-documents.json');
const data = JSON.parse(fs.readFileSync(file, 'utf8'));
const escape = s => s.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const slug = s => s.toLowerCase().replace(/<[^>]*>/g,'').replace(/[^\p{L}\p{N}_\-\s]/gu,'').replace(/\s/g,'-');
for (const doc of data.documents) {
  let markdown = doc.markdown;
  let metadata = '';
  const front = markdown.match(/^---\r?\n[\s\S]*?\r?\n---(?:\r?\n|$)/);
  if(front) { metadata = `<details class="sd-metadata"><summary>Document metadata</summary><pre>${escape(front[0])}</pre></details>`; markdown = markdown.slice(front[0].length); }
  const used = new Map(); doc.outline = [];
  const renderer = new marked.Renderer();
  renderer.heading = function(token) {
    const label = token.text.replace(/`/g,'').replace(/\*\*/g,'');
    const root = slug(label), n = used.get(root) || 0; used.set(root,n+1);
    const anchor = root + (n ? '-'+n : '');
    doc.outline.push({anchor, label, level: token.depth});
    const level = Math.min(6, token.depth + 1);
    return `<h${level} id="sd-${doc.id}-${anchor}">${this.parser.parseInline(token.tokens)}</h${level}>\n`;
  };
  renderer.link = function(token) {
    let href = token.href;
    if(!/^(?:[a-z]+:|\/\/)/i.test(href)) {
      const [relative, hash] = href.split('#');
      const targetPath = relative ? path.posix.normalize(path.posix.join(path.posix.dirname(doc.path), relative)) : doc.path;
      const target = data.documents.find(d=>d.path===targetPath);
      if(target) href = `?doc=${target.id}${hash?'&dsection='+encodeURIComponent(hash):''}#satellite-reader`;
      else href = path.posix.normalize(path.posix.join(path.posix.dirname(doc.source),relative))+(hash?'#'+hash:'');
    }
    return `<a href="${escape(href)}"${token.title?' title="'+escape(token.title)+'"':''}>${this.parser.parseInline(token.tokens)}</a>`;
  };
  doc.html = metadata + marked.parse(markdown,{gfm:true,renderer});
  // Explicit Markdown HTML anchors remain addressable inside the reader.
  doc.html = doc.html.replace(/<a id="([^"]+)"/g,(_,id)=>`<a id="sd-${doc.id}-${id}"`);
  if(/<(?:script|iframe|object)\b|\son\w+\s*=/i.test(doc.html)) throw new Error('Unexpected executable HTML in '+doc.path);
}
data.renderer = 'marked (GitHub-flavoured Markdown); headings and local document links adapted for this reader';
fs.writeFileSync(file,JSON.stringify(data,null,2)+'\n');
console.log('Rendered complete Markdown for '+data.documents.length+' documents.');
