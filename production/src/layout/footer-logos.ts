import type { Party } from "../domain/models";
export interface LogoSize { width: number; height: number }
/** Centre an image inside a padded cell without cropping or stretching. */
export function containLogo(size: LogoSize, bounds: { left: number; top: number; width: number; height: number }) {
  const scale = Math.min(bounds.width / size.width, bounds.height / size.height);
  const width = size.width * scale, height = size.height * scale;
  return { left: bounds.left + (bounds.width - width) / 2,
    top: bounds.top + (bounds.height - height) / 2, width, height };
}
/** Decode stored, validated image data before staging any slide changes. */
export async function readLogoSizes(parties: Party[]): Promise<Record<string, LogoSize>> {
  const sizes: Record<string, LogoSize> = {};
  await Promise.all(parties.filter(p => p.logoDataUrl).map(party => new Promise<void>((resolve, reject) => {
    const image = new Image();
    const timer = setTimeout(() => reject(new Error(`Could not read logo for ${party.name}. Replace it in Settings and retry.`)), 10000);
    image.onload = () => {
      clearTimeout(timer);
      if (!image.naturalWidth || !image.naturalHeight) { reject(new Error(`Invalid logo dimensions for ${party.name}.`)); return; }
      sizes[party.id] = {width:image.naturalWidth,height:image.naturalHeight}; resolve();
    };
    image.onerror = () => { clearTimeout(timer); reject(new Error(`Could not read logo for ${party.name}. Replace it in Settings and retry.`)); };
    image.src = party.logoDataUrl!;
  })));
  return sizes;
}
/** One compact, right-aligned row; scale width and height together, never stretch. */
export function footerLogoBoxes(parties: Party[], sizes: Record<string, LogoSize>, width: number, height: number) {
  const maxHeight = height * .043, gap = width * .008, available = width * .54;
  const natural = parties.map(p => sizes[p.id] ? maxHeight * sizes[p.id]!.width / sizes[p.id]!.height : width * .075);
  const total = natural.reduce((a,b)=>a+b,0) + Math.max(0,parties.length-1)*gap;
  const scale = total > available ? available / total : 1;
  let left = width * .99 - total * scale;
  return parties.map((party,index)=>{
    const box={party,left,top:height*.9665-maxHeight*scale/2,width:natural[index]!*scale,height:maxHeight*scale};
    left += box.width + gap*scale; return box;
  });
}
