import fs from 'node:fs';
import { DOMParser } from '@xmldom/xmldom';
import { describe, expect, it } from 'vitest';
import { importDocx } from '../src/formats/docx';

const REAL_DOCX = process.env.QUIRE_TEST_DOCX ?? '';

describe.runIf(Boolean(REAL_DOCX) && fs.existsSync(REAL_DOCX))('real document (QUIRE_TEST_DOCX)', () => {
  it('has nothing Quire would lose (opens normally, not protected)', async () => {
    const res = await importDocx(new Uint8Array(fs.readFileSync(REAL_DOCX)), (x) => new DOMParser().parseFromString(x, 'text/xml') as unknown as Document);
    expect(res.lossy).toEqual([]);
  });
});

import { lossyFeatures } from '../src/formats/docx/fidelity';

describe('lossyFeatures', () => {
  const body = (inner: string) => `<w:document><w:body>${inner}</w:body></w:document>`;
  const doc = (raws: string[] = []) => ({ type: 'doc', content: raws.map((xml) => ({ type: 'docxRaw', attrs: { xml, kind: 'table', text: '' } })) });

  it('flags content Quire would drop', () => {
    expect(lossyFeatures(body('<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'), doc())).toEqual(['footnotes']);
    expect(lossyFeatures(body('<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p>'), doc())).toEqual(['tracked changes']);
    expect(lossyFeatures(body('<w:p><w:commentRangeStart w:id="0"/></w:p>'), doc())).toEqual(['comments']);
    expect(lossyFeatures(body('<w:p><w:pPr><w:sectPr/></w:pPr></w:p>'), doc())).toEqual(['section breaks']);
    expect(lossyFeatures(body('<w:p><w:r><w:instrText> TOC \\o </w:instrText></w:r></w:p>'), doc())).toEqual(['fields']);
  });

  it('ignores what is kept verbatim, PAGE fields and the final section properties', () => {
    const table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:drawing/></w:r></w:p></w:tc></w:tr></w:tbl>';
    expect(lossyFeatures(body(`${table}<w:p><w:r><w:instrText> PAGE </w:instrText></w:r></w:p><w:sectPr/>`), doc([table]))).toEqual([]);
  });
});
