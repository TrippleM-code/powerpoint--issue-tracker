import {afterEach,describe,expect,it,vi} from "vitest";
import {loadTemplate,saveTemplate,TEMPLATE_KEY,normalizeTemplate} from "../../src/services/template-settings";
afterEach(()=>vi.unstubAllGlobals());
describe("template settings",()=>{
 it("defaults legacy documents to Standard and normalizes choices",()=>{
 expect(normalizeTemplate(undefined).format).toBe("standard");
 expect(normalizeTemplate({titles:[" A ","A",null,{},"B"]}).titles).toEqual(["A","B"]);
 });
 it("saves title choices separately from existing settings",async()=>{
 const data=new Map<string,unknown>();
 vi.stubGlobal("Office",{AsyncResultStatus:{Succeeded:"ok"},context:{document:{settings:{get:(k:string)=>data.get(k),set:(k:string,v:unknown)=>data.set(k,v),saveAsync:(cb:any)=>cb({status:"ok"})}}}});
 await saveTemplate({...normalizeTemplate(undefined),format:"ice",title:"A"});await saveTemplate({...loadTemplate(),title:"B"});
 expect(loadTemplate()).toMatchObject({title:"B",titles:["A","B"]});expect([...data.keys()]).toEqual([TEMPLATE_KEY]);
 });
 it("rolls back a failed save",async()=>{
 let value:unknown={format:"standard"};
 vi.stubGlobal("Office",{AsyncResultStatus:{Succeeded:"ok"},context:{document:{settings:{get:()=>value,set:(_k:string,v:unknown)=>value=v,saveAsync:(cb:any)=>cb({status:"failed",error:{message:"Disk error"}})}}}});
 await expect(saveTemplate({...normalizeTemplate(undefined),format:"ice",title:"A"})).rejects.toThrow("Disk error");expect(value).toEqual({format:"standard"});
 });
});
