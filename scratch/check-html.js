const fs = require('fs');
const path = require('path');

const htmlPath = path.join(__dirname, '..', 'docs', 'index.html');
const html = fs.readFileSync(htmlPath, 'utf8');

// A simple regex-based HTML tag balancer to check for unclosed structural tags
const stack = [];
const tagRegex = /<\/?([a-zA-Z0-9:-]+)(?:\s+[^>]*)?>/g;

let match;
let lineNum = 1;
let lastIndex = 0;

// Count lines helper
function getLineNumber(index) {
  return html.substring(0, index).split('\n').length;
}

while ((match = tagRegex.exec(html)) !== null) {
  const fullTag = match[0];
  const tagName = match[1].toLowerCase();
  const isClose = fullTag.startsWith('</');
  const isSelfClosing = fullTag.endsWith('/>') || ['img', 'br', 'hr', 'meta', 'link', 'input', 'defs', 'rect', 'circle', 'path', 'stop', 'lineargradient', 'svg'].includes(tagName);

  if (isSelfClosing) continue;

  const currentLine = getLineNumber(match.index);

  if (isClose) {
    if (stack.length === 0) {
      console.log(`Error: Close tag </${tagName}> at line ${currentLine} has no open tag`);
    } else {
      const top = stack.pop();
      if (top.name !== tagName) {
        // Only report major structural mismatches to avoid minor noise
        if (['div', 'section', 'main', 'aside', 'ul', 'li'].includes(tagName) || ['div', 'section', 'main', 'aside', 'ul', 'li'].includes(top.name)) {
          console.log(`Warning: Mismatched close tag </${tagName}> at line ${currentLine}. Expected </${top.name}> (opened at line ${top.line})`);
          // Put top back so we can try to recover
          stack.push(top);
        }
      }
    }
  } else {
    stack.push({ name: tagName, line: currentLine });
  }
}

if (stack.length > 0) {
  console.log('\nUnclosed tags remaining at end of file:');
  stack.forEach(t => {
    if (['div', 'section', 'main', 'aside', 'ul', 'li'].includes(t.name)) {
      console.log(`- <${t.name}> opened at line ${t.line}`);
    }
  });
} else {
  console.log('No structural mismatched or unclosed tags found!');
}
