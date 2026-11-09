import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { createServer } from 'node:http';
import { promisify } from 'node:util';
import test from 'node:test';
import { JSDOM } from '../lib/jsdom.js';

test('bundled DOM preserves XML namespaces, international URLs and default styles', () => {
  const xml = new JSDOM('<a:root xmlns:a="urn:ppt"><a:text>可编辑</a:text></a:root>', { contentType: 'text/xml' });
  const text = xml.window.document.getElementsByTagNameNS('urn:ppt', 'text')[0];
  assert.equal(text.textContent, '可编辑');
  assert.match(new xml.window.XMLSerializer().serializeToString(text), /xmlns:a="urn:ppt"/);
  xml.window.close();
  const html = new JSDOM('<p>text</p>', { url: 'https://例子.测试/' });
  assert.equal(html.window.location.hostname, 'xn--fsqu00a.xn--0zwm56d');
  assert.equal(html.window.getComputedStyle(html.window.document.querySelector('p')).display, 'block');
  html.window.close();
});

test('bundled DOM synchronous XHR starts its packaged worker', { timeout: 30000 }, async t => {
  const server = createServer((_req, res) => { res.end('worker response'); });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  t.after(() => new Promise(resolve => server.close(resolve)));
  const url = `http://127.0.0.1:${server.address().port}/`;
  const { stdout } = await promisify(execFile)(process.execPath, ['--input-type=module', '-e', `
    import { JSDOM } from ${JSON.stringify(new URL('../lib/jsdom.js', import.meta.url).href)};
    const dom = new JSDOM('', { url: ${JSON.stringify(url)} });
    const xhr = new dom.window.XMLHttpRequest();
    xhr.open('GET', ${JSON.stringify(url)}, false);
    xhr.send();
    console.log(xhr.status, xhr.responseText);
    dom.window.close();
  `], { timeout: 25000 });
  assert.equal(stdout.trim(), '200 worker response');
});
