// Reads the cover page (R1.htm) of each company's latest 10-Q/10-K to get shares outstanding by class.
const fs=require('fs'),path=require('path');const OUT=path.join(__dirname,'..','data','raw','sec_cover');
const UA={headers:{'User-Agent':'fin-workshop research admin@fin-workshop.local'}};const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const CIKS={'BRK.B':1067983,V:1403161,ARES:1176948,HSY:47111,STZ:16918,ERIE:922621,TSN:100493,KVUE:1944048,EME:105634,VYLR:2128626};
(async()=>{for(const [sym,cik] of Object.entries(CIKS)){
  const sub=await (await fetch(`https://data.sec.gov/submissions/CIK${String(cik).padStart(10,'0')}.json`,UA)).json();await sleep(150);
  const r=sub.filings.recent;let k=r.form.findIndex(f=>f==='10-Q'||f==='10-K');
  if(k<0){k=r.form.findIndex(f=>/^(S-1|424B4|10-12B|8-K)/.test(f));}
  const acc=r.accessionNumber[k],form=r.form[k],date=r.filingDate[k];
  const url=`https://www.sec.gov/Archives/edgar/data/${cik}/${acc.replace(/-/g,'')}/R1.htm`;
  const res=await fetch(url,UA);await sleep(150);const html=res.ok?await res.text():'';
  const text=html.replace(/<[^>]+>/g,'|').replace(/&nbsp;|&#160;/g,' ').replace(/\|\s*\|+/g,'|').replace(/\s+/g,' ');
  const hits=[...text.matchAll(/([^|]{0,120})\|\s*Entity Common Stock, Shares Outstanding\s*\|([^A-Za-z]{0,200})/g)].map(m=>m[0]);
  // Also capture class-labelled rows from the header area
  const rows=[...text.matchAll(/Shares Outstanding[^|]*\|([\d,| ]+)/g)].map(m=>m[1].split('|').map(s=>s.trim()).filter(Boolean));
  const cls=[...text.matchAll(/\|\s*(Class [A-Z0-9-]+[^|]{0,40}|Common Stock[^|]{0,60}|Series [A-Z][^|]{0,40})\s*\|/g)].map(m=>m[1].trim());
  fs.writeFileSync(path.join(OUT,sym+'.html'),html);
  console.log(sym,form,date,res.status,'| classes:',[...new Set(cls)].slice(0,8).join(' ; '),'| shares rows:',JSON.stringify(rows).slice(0,300));
}})();
