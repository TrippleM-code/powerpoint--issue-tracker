import { footerLogoBoxes, type LogoSize } from "./footer-logos";
import type { Issue, Party } from "../domain/models";
import type { TemplateSettings } from "../services/template-settings";
import { TAGS } from "../storage/tag-names";

export interface PageSize { width: number; height: number }
export interface ReferenceBox { role: string; left: number; top: number; width: number; height: number }
export function referenceBoxes(ice: boolean, size: PageSize): ReferenceBox[] {
  return ["SOI", "NOV", "COI/CVI"].map((role, index) => ice ? {
    role, left: size.width * (.745 + index * .083), top: size.height * .051,
    width: size.width * .079, height: size.height * .038,
  } : { role, left: 776 + index * (160 / 3 + 4), top: 32, width: 160 / 3, height: 28 });
}
export interface IceDrawing {
  text: (slide: any, text: string, l: number, t: number, w: number, h: number, options?: {size?: number; bold?: boolean; color?: string}) => any;
  rect: (slide: any, l: number, t: number, w: number, h: number, fill: string, line: string) => any;
  tag: (shape: any, role: string) => void;
  colors: (status: string) => {fill: string; text: string};
}
/** Transparent column bodies leave user-owned drawings visible on refresh. */
export function renderIceIssue(slide: any, issue: Issue, template: TemplateSettings, parties: Party[], size: PageSize, existingRoles: Set<string>, draw: IceDrawing, logoSizes: Record<string, LogoSize> = {}): void {
  const W = size.width, H = size.height, font = W / 960;
  const navy = "#083479", grid = "#17344D";
  let sequence = 0;
  const text = (value: string, x: number, y: number, w: number, h: number, point = 10, color = grid, bold = false) => {
    const shape = draw.text(slide, value, x*W, y*H, w*W, h*H, {size:point*font,color,bold});
    draw.tag(shape, `ICE_TEXT_${sequence++}`); return shape;
  };
  const rect = (x: number, y: number, w: number, h: number, fill = navy) => {
    const shape = draw.rect(slide,x*W,y*H,w*W,h*H,fill,grid);
    draw.tag(shape, `ICE_BOX_${sequence++}`); return shape;
  };
  rect(.01,.016,.207,.073);
  text(`ICE MEETING ${template.meetingNumber || "—"} AGENDA\nDATE : ${template.meetingDate || "—"}`, .018,.022,.191,.061,10,"#FFFFFF",true);
  text(template.title,.224,.017,.51,.035,11,grid,true);
  text(`COORDINATION AREA : ${issue.areaCode}   |   ${issue.roomSpace}`, .224,.054,.51,.031,10);
  referenceBoxes(true,size).forEach(box => {
    const heading = draw.text(slide,box.role,box.left,.016*H,box.width,.035*H,{size:9*font,bold:true,color:"#FFFFFF"});
    heading.fill.setSolidColor("#626D79"); draw.tag(heading,`REFERENCE_HEADER_${box.role}`);
    if (!existingRoles.has(box.role)) {
      const field = draw.text(slide,"",box.left,box.top,box.width,box.height,{size:9*font});
      field.fill.setSolidColor("#FFFFFF"); field.lineFormat.color=grid; field.lineFormat.weight=.7;
      field.tags.add(TAGS.referenceRole,box.role);
    }
  });
  const xs=[.01,.079,.455,.831,.917,.996];
  ["S/N","DESCRIPTION","SOLUTION","ACTION BY","STATUS"].forEach((label,i)=>{
    const x=xs[i]!, width=xs[i+1]!-x;
    rect(x,.106,width,.032); text(label,x+.003,.108,width-.006,.026,8,"#FFFFFF",true);
    const border=rect(x,.138,width,.792,"#FFFFFF"); border.fill.transparency=1;
  });
  text(issue.id,.013,.147,.063,.045,10,grid,true);
  text(issue.description,.091,.15,.35,.22,11);
  // Description below 37% is reserved for manual reference images/annotations.
  const visible=issue.actions.slice(0,10), overflow=issue.actions.length-visible.length;
  const rowHeight=(.792-(overflow?.05:0))/Math.max(visible.length,1);
  if (!visible.length) text("No actions added.",.467,.15,.35,.04);
  visible.forEach((action,i)=>{
    const y=.138+i*rowHeight, colors=draw.colors(action.status);
    if(i) rect(.455,y,.541,.001,"#C7D6E5");
    text(action.required,.467,y+.01,.352,rowHeight-.02,10);
    text(action.party,.835,y+.01,.078,rowHeight-.02,9,grid,true);
    rect(.917,y,.079,rowHeight,colors.fill);
    text(action.status,.92,y+.01,.073,rowHeight-.02,9,colors.text,true);
  });
  if(overflow) text(`+ ${overflow} more action(s). See Actions tab / full Action Register.`,.467,.885,.52,.038,8,grid,true);
  text(`Created ${new Date(issue.createdAt).toLocaleString()}\nUpdated ${issue.updatedAt ? new Date(issue.updatedAt).toLocaleString() : "—"}`,.01,.94,.42,.047,7);
  footerLogoBoxes(parties,logoSizes,W,H).forEach(box=>{
    if(box.party.logoDataUrl){
      const logo=draw.rect(slide,box.left,box.top,box.width,box.height,"#FFFFFF",grid);
      draw.tag(logo,`ICE_LOGO_${box.party.id}`); logo.lineFormat.transparency=1;
      logo.fill.setImage(box.party.logoDataUrl.slice(box.party.logoDataUrl.indexOf(",")+1));
    } else text(box.party.name,box.left/W,box.top/H,box.width/W,box.height/H,8,grid,true);
  });
}