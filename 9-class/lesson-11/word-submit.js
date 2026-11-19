/* Word (.docx) export and submission for 9th grade.
 * Creates a standards-based OOXML document in a stored ZIP archive locally.
 * No third-party script or external conversion service is used.
 */
(function(){
"use strict";
const RECIPIENT="work.inform@chern12.ukr.education";
const MIME="application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const utf8=new TextEncoder();
const ns={
 w:"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
 r:"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
 wp:"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
 a:"http://schemas.openxmlformats.org/drawingml/2006/main",
 pic:"http://schemas.openxmlformats.org/drawingml/2006/picture"
};
const el=id=>document.getElementById(id);
const xesc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&apos;"}[c])).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,"");
const clean=s=>String(s??"").replace(/\s+/g," ").trim();
function xmlwrap(node,content){
 const tag=node.trim().split(/\s+/)[0];
 return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><'+node+'>'+content+'</'+tag+'>';
}
function zipStored(entries){
 const crcTable=zipStored.crcTable||(zipStored.crcTable=Array.from({length:256},(_,i)=>{let c=i;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;return c>>>0;}));
 function crc32(buf){let c=0xFFFFFFFF;for(let i=0;i<buf.length;i++)c=crcTable[(c^buf[i])&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0;}
 const chunks=[],centers=[];let offset=0,centralLength=0;
 for(const [filename,body] of entries){
  const name=utf8.encode(filename),data=typeof body==="string"?utf8.encode(body):body;
  if(!(data instanceof Uint8Array))throw new Error("Неправильний формат даних у документі.");
  const crc=crc32(data),local=new Uint8Array(30+name.length),head=new DataView(local.buffer);
  head.setUint32(0,0x04034b50,true);head.setUint16(4,20,true);head.setUint16(6,0,true);
  head.setUint16(8,0,true);head.setUint32(14,crc,true);
  head.setUint32(18,data.length,true);head.setUint32(22,data.length,true);
  head.setUint16(26,name.length,true);local.set(name,30);
  chunks.push(local,data);
  const center=new Uint8Array(46+name.length),v=new DataView(center.buffer);
  v.setUint32(0,0x02014b50,true);v.setUint16(4,20,true);v.setUint16(6,20,true);
  v.setUint16(10,0,true);v.setUint32(16,crc,true);v.setUint32(20,data.length,true);
  v.setUint32(24,data.length,true);v.setUint16(28,name.length,true);
  v.setUint32(42,offset,true);center.set(name,46);centers.push(center);
  offset+=local.length+data.length;centralLength+=center.length;
 }
 const tail=new Uint8Array(22),d=new DataView(tail.buffer);
 d.setUint32(0,0x06054b50,true);d.setUint16(8,entries.length,true);
 d.setUint16(10,entries.length,true);d.setUint32(12,centralLength,true);d.setUint32(16,offset,true);
 return new Blob([...chunks,...centers,tail],{type:MIME});
}
function dataBytes(url){
 const split=String(url).indexOf(",");
 if(split<0)throw new Error("Неправильне зображення.");
 const txt=atob(String(url).slice(split+1)),out=new Uint8Array(txt.length);
 for(let i=0;i<txt.length;i++)out[i]=txt.charCodeAt(i);
 return out;
}
function wordFile(student,klass,score){
 const rels=[
  {id:"rId1",type:"styles",target:"styles.xml"},
  {id:"rId2",type:"numbering",target:"numbering.xml"},
  {id:"rId3",type:"header",target:"header1.xml"},
  {id:"rId4",type:"footer",target:"footer1.xml"}
 ];
 let counter=5,imageNumber=0,bookmark=1;
 const media=[],contentImages=new Set();
 const anchorMap=new Map();
 function rel(type,target,external=false){
  const id="rId"+counter++;
  rels.push({id,type,target,external});return id;
 }
 const pxToTwip=px=>Math.round(px*15);
 function colorHex(color){
  const s=String(color||"").trim().toLowerCase();
  if(!s||s==="transparent"||s==="none")return null;
  if(/^#[0-9a-f]{6}$/.test(s))return s.slice(1).toUpperCase();
  if(/^#[0-9a-f]{8}$/.test(s))return parseInt(s.slice(7,9),16)===0?null:s.slice(1,7).toUpperCase();
  if(/^#[0-9a-f]{3}$/.test(s))return s.slice(1).split("").map(x=>x+x).join("").toUpperCase();
  // The transparent computed browser default is rgba(0, 0, 0, 0).
  // Ignoring the alpha channel used to export it as BLACK Word shading.
  const m=s.match(/^rgba?\(\s*(\d{1,3})\s*[, ]+\s*(\d{1,3})\s*[, ]+\s*(\d{1,3})(?:\s*[,/]\s*([\d.]+))?\s*\)$/);
  if(!m)return null;
  if(s.startsWith("rgba")&&(m[4]===undefined||Number(m[4])<=0))return null;
  return m.slice(1,4).map(x=>Math.max(0,Math.min(255,Number(x))).toString(16).padStart(2,"0")).join("").toUpperCase();
 }
 function textRun(node){
  if(node.nodeType===Node.TEXT_NODE){
   if(!node.nodeValue)return "";
   const parent=node.parentElement||editor,c=getComputedStyle(parent);
   const fontName=(c.fontFamily.split(",")[0]||"Times New Roman").trim().replace(/["']/g,"").slice(0,65);
   const sz=Math.max(14,Math.min(60,Math.round(parseFloat(c.fontSize||"16")*1.5)));
   const weight=Number(c.fontWeight);
   let rpr='<w:rFonts w:ascii="'+xesc(fontName)+'" w:hAnsi="'+xesc(fontName)+'"/><w:sz w:val="'+sz+'"/>';
   if(c.fontWeight==="bold"||weight>=600)rpr+='<w:b/>';
   if(c.fontStyle==="italic")rpr+='<w:i/>';
   if(c.textDecorationLine.includes("underline"))rpr+='<w:u w:val="single"/>';
   const fg=colorHex(c.color);if(fg)rpr+='<w:color w:val="'+fg+'"/>';
   for(let p=parent;p&&p!==editor;p=p.parentElement){
    const bg=colorHex(p.style?.backgroundColor); // export only intentional user-applied marker, never transparent computed styles
    if(bg&&bg!=="FFFFFF"&&bg!=="F7FAFF"){rpr+='<w:shd w:val="clear" w:color="auto" w:fill="'+bg+'"/>';break;}
   }
   return '<w:r><w:rPr>'+rpr+'</w:rPr><w:t xml:space="preserve">'+xesc(node.nodeValue)+'</w:t></w:r>';
  }
  if(node.nodeType!==Node.ELEMENT_NODE)return "";
  const tag=node.tagName;
  if(tag==="BR")return '<w:r><w:br/></w:r>';
  if(tag==="IMG")return imageRun(node);
  if(tag==="A"){
   const url=safeUrl(node.getAttribute("href"));
   const content=Array.from(node.childNodes).map(textRun).join("")||'<w:r><w:t>'+xesc(node.textContent)+'</w:t></w:r>';
   if(!url)return content;
   const id=rel("hyperlink",url,true);
   return '<w:hyperlink r:id="'+id+'" w:history="1">'+content+'</w:hyperlink>';
  }
  if(["UL","OL","TABLE"].includes(tag))return "";
  return Array.from(node.childNodes).map(textRun).join("");
 }
 function imageRun(img){
  const src=img.getAttribute("src")||"";
  const kind=(src.match(/^data:image\/(png|jpeg|webp);base64,/i)||[])[1]?.toLowerCase();
  if(!kind)return "";
  let buf;try{buf=dataBytes(src);}catch{return "";}
  if(buf.byteLength>3500000)return "";
  const ext=kind==="jpeg"?"jpg":kind;
  const name="image"+(++imageNumber)+"."+ext;
  media.push(["word/media/"+name,buf]);contentImages.add(ext);
  const id=rel("image","media/"+name);
  const wpId=imageNumber+100,ratio=(img.naturalHeight&&img.naturalWidth)?img.naturalHeight/img.naturalWidth:.5;
  const width=380*9525,height=Math.max(60,Math.round(380*ratio))*9525;
  return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'+
   '<wp:extent cx="'+width+'" cy="'+height+'"/><wp:docPr id="'+wpId+'" name="Ілюстрація '+imageNumber+'"/>'+
   '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'+
   '<pic:pic><pic:nvPicPr><pic:cNvPr id="'+wpId+'" name="'+xesc(name)+'"/><pic:cNvPicPr/></pic:nvPicPr>'+
   '<pic:blipFill><a:blip r:embed="'+id+'"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'+
   '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+width+'" cy="'+height+'"/></a:xfrm>'+
   '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'+
   '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
 }
 function para(node,opts={}){
  const tag=node?.tagName||"P",cs=node&&node.isConnected?getComputedStyle(node):null;
  let ppr="";
  if(tag==="H1"||tag==="H2")ppr+='<w:pStyle w:val="'+(tag==="H1"?"Heading1":"Heading2")+'"/>';
  if(opts.title)ppr+='<w:pStyle w:val="Title"/>';
  if(opts.num){
   ppr+='<w:numPr><w:ilvl w:val="'+Math.min(8,opts.level||0)+'"/><w:numId w:val="'+(opts.num==="ol"?"2":"1")+'"/></w:numPr>';
  }
  if(cs){
   const alignment={left:"left",center:"center",right:"right",justify:"both"}[cs.textAlign];
   if(alignment)ppr+='<w:jc w:val="'+alignment+'"/>';
   const first=parseFloat(cs.textIndent),left=parseFloat(cs.marginLeft);
   if((Number.isFinite(first)&&first>0)||(Number.isFinite(left)&&left>0)){
    ppr+='<w:ind'+(left>0?' w:left="'+pxToTwip(left)+'"':'')+(first>0?' w:firstLine="'+pxToTwip(first)+'"':'')+'/>';
   }
   if(["1","1.15","1.5","2"].includes(node.style.lineHeight))ppr+='<w:spacing w:line="'+Math.round(Number(node.style.lineHeight)*240)+'" w:lineRule="auto"/>';
  }
  let texts="";
  if(node){
   const visible=Array.from(node.childNodes).filter(n=>!["OL","UL","TABLE"].includes(n.nodeName));
   texts=visible.map(textRun).join("");
  }else texts='<w:r><w:t xml:space="preserve">'+xesc(opts.value||"")+'</w:t></w:r>';
  if(!texts)texts="<w:r><w:t></w:t></w:r>";
  let bookmarks="";
  if(tag==="H1"||tag==="H2"){
   const id=bookmark++,label="section"+id;
   anchorMap.set(node,label);
   bookmarks='<w:bookmarkStart w:id="'+id+'" w:name="'+label+'"/>';
   texts=bookmarks+texts+'<w:bookmarkEnd w:id="'+id+'"/>';
  }
  return '<w:p><w:pPr>'+ppr+'</w:pPr>'+texts+'</w:p>';
 }
 function list(node,level=0){
  const parts=[],type=node.tagName==="OL"?"ol":"ul";
  for(const li of node.children)if(li.tagName==="LI"){
   parts.push(para(li,{num:type,level}));
   for(const child of li.children)if(["UL","OL"].includes(child.tagName))parts.push(list(child,level+1));
  }
  return parts.join("");
 }
 function table(t){
  let s='<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>';
  for(const b of ["top","left","bottom","right","insideH","insideV"])s+='<w:'+b+' w:val="single" w:sz="6" w:color="9FB2CB"/>';
  s+='</w:tblBorders></w:tblPr><w:tblGrid>';
  const count=Math.min(10,Math.max(1,t.querySelector("tr")?.children.length||3));
  for(let i=0;i<count;i++)s+='<w:gridCol w:w="'+Math.floor(9000/count)+'"/>';
  s+='</w:tblGrid>';
  for(const tr of t.querySelectorAll("tr")){
   s+='<w:tr>';
   for(const cell of tr.children)if(["TD","TH"].includes(cell.tagName)){
    const fill=cell.tagName==="TH"?(t.classList.contains("mint")?"D9F0E8":t.classList.contains("gray")?"E9EAF0":"DCE8FA"):null;
    s+='<w:tc><w:tcPr><w:tcW w:w="'+Math.floor(9000/count)+'" w:type="dxa"/>'+(fill?'<w:shd w:fill="'+fill+'"/>':'')+'</w:tcPr>';
    s+=para(cell)+'</w:tc>';
   }
   s+='</w:tr>';
  }
  return s+'</w:tbl>';
 }
 function blocks(root){
  let output="";
  for(const node of root.children){
   switch(node.tagName){
    case "UL":case "OL":output+=list(node);break;
    case "TABLE":output+=table(node);break;
    case "HR":output+='<w:p><w:r><w:br w:type="page"/></w:r></w:p>';break;
    case "IMG":output+=para(node.parentElement===root?null:node,{value:""});if(node.parentElement===root)output+='<w:p>'+imageRun(node)+'</w:p>';break;
    case "P":case "H1":case "H2":case "DIV":
     if(node.querySelector("table,ul,ol")){if(clean(node.textContent))output+=para(node);for(const nested of node.children)if(nested.tagName==="TABLE")output+=table(nested);else if(["UL","OL"].includes(nested.tagName))output+=list(nested);}
     else output+=para(node);break;
    default:if(clean(node.textContent))output+=para(node);break;
   }
  }
  return output;
 }
 const title=el("docName").value.trim()||"Документ";
 const output=[para(null,{title:true,value:title})];
 output.push(para(null,{value:"Учень: "+student+"    Клас: "+klass+"    Виконано: "+score+" із "+TASKS.length}));
 output.push(para(null,{value:"Інформатика, 9 клас. Практична робота № 3."}));
 // Convert the document body first so headings receive unique bookmark IDs.
 const body=blocks(editor);
 if(tocActive){
  const items=[...editor.querySelectorAll("h1,h2")].filter(h=>clean(h.textContent));
  if(items.length){
   output.push('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Зміст</w:t></w:r></w:p>');
   for(const h of items){
    const dest=anchorMap.get(h);
    if(dest){
     const margin=h.tagName==="H2"?'<w:ind w:left="560"/>':"";
     output.push('<w:p><w:pPr>'+margin+'</w:pPr><w:hyperlink w:anchor="'+dest+'"><w:r><w:rPr><w:color w:val="25529A"/><w:u w:val="single"/></w:rPr><w:t>'+xesc(clean(h.textContent))+'</w:t></w:r></w:hyperlink></w:p>');
    }
   }
  }
 }
 output.push(body);
 const relsXml=xmlwrap('Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"',
  rels.map(r=>'<Relationship Id="'+r.id+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'+r.type+'" Target="'+xesc(r.target)+'"'+(r.external?' TargetMode="External"':'')+'/>').join(""));
 const headerText=clean(el("headerText").value);
 const footerText=clean(el("footerText").value);
 const header=xmlwrap('w:hdr xmlns:w="'+ns.w+'"',para(null,{value:headerText}));
 const footer=xmlwrap('w:ftr xmlns:w="'+ns.w+'"',para(null,{value:footerText}));
 const docBody=output.join("")+
 '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId4"/>'+
 '<w:pgSz w:w="'+(orientation==="landscape"?"16838":"11906")+'" w:h="'+(orientation==="landscape"?"11906":"16838")+'"/>'+
 '<w:pgMar w:top="1200" w:right="1200" w:bottom="1200" w:left="1200" w:header="550" w:footer="550" w:gutter="0"/></w:sectPr>';
 const doc=xmlwrap('w:document xmlns:w="'+ns.w+'" xmlns:r="'+ns.r+'" xmlns:wp="'+ns.wp+'" xmlns:a="'+ns.a+'" xmlns:pic="'+ns.pic+'"','<w:body>'+docBody+'</w:body>');
 const styles=xmlwrap('w:styles xmlns:w="'+ns.w+'"',
  '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/>'+
  '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:style>'+
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:spacing w:after="260"/></w:pPr><w:rPr><w:b/><w:sz w:val="38"/></w:rPr></w:style>'+
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="280" w:after="140"/></w:pPr><w:rPr><w:b/><w:color w:val="244A86"/><w:sz w:val="32"/></w:rPr></w:style>'+
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:outlineLvl w:val="1"/><w:spacing w:before="200" w:after="100"/></w:pPr><w:rPr><w:b/><w:color w:val="365C92"/><w:sz w:val="28"/></w:rPr></w:style>');
 function level(kind,depth){
  const bullet=kind==="bullet",symbols=["•","◦","▪"];
  return '<w:lvl w:ilvl="'+depth+'"><w:start w:val="1"/><w:numFmt w:val="'+(bullet?"bullet":"decimal")+'"/>'+
   '<w:lvlText w:val="'+(bullet?symbols[depth%3]:(depth===0?"%1.":"%"+(depth+1)+"."))+'"/>'+
   '<w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="'+((depth+1)*640)+'"/></w:tabs>'+
   '<w:ind w:left="'+((depth+1)*640)+'" w:hanging="300"/></w:pPr></w:lvl>';
 }
 const numbering=xmlwrap('w:numbering xmlns:w="'+ns.w+'"',
  '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>'+
  [0,1,2,3].map(n=>level("bullet",n)).join("")+'</w:abstractNum>'+
  '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/>'+
  [0,1,2,3].map(n=>level("decimal",n)).join("")+'</w:abstractNum>'+
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>');
 const types=[
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
  '<Default Extension="xml" ContentType="application/xml"/>',
  ...[...contentImages].map(ext=>'<Default Extension="'+ext+'" ContentType="image/'+(ext==="jpg"?"jpeg":ext)+'"/>'),
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
  '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
  '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
  '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
 ];
 const contentTypes=xmlwrap('Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"',types.join(""));
 const rootRels=xmlwrap('Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"',
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>');
 const files=[
  ["[Content_Types].xml",contentTypes],
  ["_rels/.rels",rootRels],
  ["word/document.xml",doc],
  ["word/_rels/document.xml.rels",relsXml],
  ["word/styles.xml",styles],["word/numbering.xml",numbering],
  ["word/header1.xml",header],["word/footer1.xml",footer],...media
 ];
 const blob=zipStored(files);
 if(blob.size>9500000)throw new Error("Файл перевищує ліміт 10 МБ для надсилання через FormSubmit.");
 return blob;
}
function scoreNow(){
 let n=0;for(const t of TASKS){try{if(t[3]())n++;}catch{}}
 return n;
}
function details(required){
 const student=clean(el("pupilName").value).slice(0,70);
 const klass=el("pupilClass").value;
 if(required&&(!student||!["9-А","9-Б"].includes(klass))){
  el("submitError").textContent="Укажіть клас та ім’я учня перед надсиланням.";
  el("pupilName").focus();throw new Error("Неповні дані учня.");
 }
 return {student:student||"Учень",klass:klass||"9 клас",score:scoreNow()};
}
function filename({student,klass}){
 return ("ПР3_"+klass+"_"+student).replace(/[^\p{L}\p{N}_-]+/gu,"_").slice(0,115)+".docx";
}
function downloadBlob(blob,name){
 const url=URL.createObjectURL(blob),a=document.createElement("a");
 a.href=url;a.download=name;document.body.append(a);a.click();a.remove();
 setTimeout(()=>URL.revokeObjectURL(url),6000);
}
async function downloadWord(){
 try{const data=details(false);const blob=wordFile(data.student,data.klass,data.score);
  downloadBlob(blob,filename(data));toast("Word-документ (.docx) збережено.");
 }catch(e){el("submitError").textContent=e.message;toast("Не вдалося підготувати Word-файл.");}
}
let sending=false;
async function submitWord(){
 if(sending)return;
 const form=el("wordSendForm"),button=el("sendWord");
 el("submitError").textContent="";
 let data;
 try{data=details(true);}catch{return;}
 sending=true;button.disabled=true;button.textContent="Підготовка Word-файлу…";
 try{
  const blob=wordFile(data.student,data.klass,data.score);
  const file=new File([blob],filename(data),{type:MIME});
  const transfer=new DataTransfer();transfer.items.add(file);
  el("wordAttachment").files=transfer.files;
  if(!el("wordAttachment").files?.length)throw new Error("Браузер не дозволив додати вкладення.");
  el("formSubject").value="Інформатика · Практична №3 · "+data.klass+" · "+data.student;
  el("formProgress").value=data.score+" із "+TASKS.length;
  el("formFeedback").value="Учень: "+data.student+"; клас: "+data.klass+
   "; автоматичний прогрес: "+data.score+"/"+TASKS.length+
   "; документ: "+el("docName").value+"; вкладення: "+file.name+".";
  el("formNext").value=location.origin+location.pathname+"?submitted=1";
  // Native multipart form submission is required for FormSubmit attachments.
  // The service will redirect to its verification page on first use.
  el("submitReady").value="true";
  form.requestSubmit();
 }catch(e){
  el("submitError").textContent=e.message+" Можна окремо зберегти .docx.";
  sending=false;button.disabled=false;button.textContent="Надіслати роботу на пошту (.docx)";
 }
}
const form=el("wordSendForm");
form.addEventListener("submit",e=>{
 if(el("submitReady").value!=="true"||!el("wordAttachment").files?.length){
  e.preventDefault();el("submitError").textContent="Спочатку натисніть кнопку надсилання, щоб сформувати файл.";
 }
});
el("sendWord").addEventListener("click",submitWord);
el("saveWordTask").addEventListener("click",downloadWord);
if(new URLSearchParams(location.search).get("submitted")==="1"){
 el("submissionNotice").hidden=false;
 try{history.replaceState(null,"",location.pathname);}catch{}
}
window.wordSubmission={download:downloadWord};
})();
