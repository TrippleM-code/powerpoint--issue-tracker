import {describe,it,expect} from "vitest";
import {containLogo,footerLogoBoxes} from "../../src/layout/footer-logos";
describe("Standard header proportions",()=>{
 it.each([[400,100],[100,100],[100,400],[1200,100]])("fits %s x %s logos without stretching or cropping",(width,height)=>{
  const bounds={left:14,top:23,width:115,height:32};
  const box=containLogo({width,height},bounds);
  expect(box.width/box.height).toBeCloseTo(width/height);
  expect(box.left+box.width/2).toBeCloseTo(bounds.left+bounds.width/2);
  expect(box.top+box.height/2).toBeCloseTo(bounds.top+bounds.height/2);
  expect(box.width).toBeLessThanOrEqual(bounds.width+1e-9);
  expect(box.height).toBeLessThanOrEqual(bounds.height+1e-9);
 });
});
describe("ICE footer proportions",()=>{
 it.each([[960,540],[720,540],[1008,756]])("preserves wide and square logos on %s x %s slides",(w,h)=>{
 const parties=[{id:"wide",name:"Wide"},{id:"square",name:"Square"},{id:"strip",name:"Combined"}];
 const boxes=footerLogoBoxes(parties,{wide:{width:400,height:100},square:{width:100,height:100},strip:{width:1200,height:100}},w,h);
 boxes.forEach((b,i)=>expect(b.width/b.height).toBeCloseTo([4,1,12][i]!));
 expect(boxes[2]!.left+boxes[2]!.width).toBeCloseTo(w*.99);
 expect(boxes[0]!.left).toBeGreaterThanOrEqual(w*.45-.01);
 boxes.forEach(b=>expect(b.top+b.height).toBeLessThan(h));
 });
});
