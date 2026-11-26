/** Plain inline pictures: recognised on import, written (or patched) on export. */
import { EMU_PER_PX } from '../images';
import { attr, childByName, elementChildren, escapeXml, findAllDescendants, findDescendant } from './xml';

export interface SimplePicture {
  relId: string;
  width: number;
  height: number;
  alt: string;
}

const px = (emu: number): number => Math.round((emu / EMU_PER_PX) * 100) / 100;

/**
 * Reads a `w:drawing` that is just a picture in line with the text. Anything more (floating, cropped,
 * rotated, linked, with effects or a shape) returns null and stays an untouched kept-as-is block.
 */
export function readSimplePicture(drawing: Element): SimplePicture | null {
  const kids = elementChildren(drawing);
  const inline = kids[0];
  if (kids.length !== 1 || inline.nodeName !== 'wp:inline') return null;
  const extent = childByName(inline, 'wp:extent');
  const cx = Number(extent && attr(extent, 'cx'));
  const cy = Number(extent && attr(extent, 'cy'));
  if (!(cx > 0) || !(cy > 0)) return null;
  const graphicData = findDescendant(inline, 'a:graphicData');
  const pics = graphicData ? elementChildren(graphicData) : [];
  if (pics.length !== 1 || pics[0].nodeName !== 'pic:pic') return null;
  const pic = pics[0];
  if (findDescendant(inline, 'a:hlinkClick') || findDescendant(inline, 'a:hlinkHover')) return null;
  const blips = findAllDescendants(pic, 'a:blip');
  const relId = blips.length === 1 ? attr(blips[0], 'r:embed') : null;
  if (!relId || attr(blips[0], 'r:link') || elementChildren(blips[0]).some((e) => e.nodeName !== 'a:extLst')) return null;
  const crop = findDescendant(pic, 'a:srcRect');
  if (crop && ['l', 't', 'r', 'b'].some((k) => Number(attr(crop, k) ?? 0) !== 0)) return null;
  const xfrm = findDescendant(pic, 'a:xfrm');
  if (xfrm && ['rot', 'flipH', 'flipV'].some((k) => !['', '0', 'false'].includes(attr(xfrm, k) ?? ''))) return null;
  const geom = findDescendant(pic, 'a:prstGeom');
  if (findDescendant(pic, 'a:custGeom') || (geom && attr(geom, 'prst') !== 'rect')) return null;
  const docPr = childByName(inline, 'wp:docPr');
  return { relId, width: px(cx), height: px(cy), alt: (docPr && attr(docPr, 'descr')) ?? '' };
}

/** Verbatim `w:drawing` elements in a paragraph's source (plain pictures never nest). */
export function drawingsIn(xml: string): string[] {
  return xml.match(/<w:drawing\b[\s\S]*?<\/w:drawing>/g) ?? [];
}

/** The image relationship a verbatim drawing uses. */
export function blipRel(xml: string): string | null {
  return /<a:blip\b[^>]*?\sr:embed="([^"]+)"/.exec(xml)?.[1] ?? null;
}

/** The drawing's object id (must be unique in the document). */
export function docPrIds(xml: string): number[] {
  return [...xml.matchAll(/<wp:docPr\b[^>]*?\sid="(\d+)"/g)].map((m) => Number(m[1]));
}

/** A verbatim drawing with its size changed; untouched when the size is the same. */
export function resizeDrawing(xml: string, width: number, height: number): string {
  const cur = /<wp:extent\b[^>]*>/.exec(xml)?.[0] ?? '';
  const read = (k: string) => Number(new RegExp(`\\s${k}="(\\d+)"`).exec(cur)?.[1]);
  if (px(read('cx')) === width && px(read('cy')) === height) return xml;
  const set = (tag: string): string =>
    tag.replace(/(\scx=")\d+/, `$1${emu(width)}`).replace(/(\scy=")\d+/, `$1${emu(height)}`);
  return xml.replace(/<wp:extent\b[^>]*>/, set).replace(/<a:ext\b(?=[^>]*\scx=)[^>]*>/g, set);
}

function emu(v: number): number {
  return Math.max(1, Math.round(v * EMU_PER_PX));
}

/** A new inline picture. Namespaces are declared on the spot, so it fits any document root. */
export function newDrawing(relId: string, id: number, width: number, height: number, alt: string): string {
  const size = `cx="${emu(width)}" cy="${emu(height)}"`;
  const descr = alt ? ` descr="${escapeXml(alt)}"` : '';
  return (
    '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    `distT="0" distB="0" distL="0" distR="0"><wp:extent ${size}/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="Picture ${id}"${descr}/>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"${descr}/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${relId}"/>` +
    '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext ${size}/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
  );
}
