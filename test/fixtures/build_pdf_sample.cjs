// Builds pdf-styled-sample.pdf (manual check): mid-sentence colors, links with
// underlines, an image with text in it, and white text over a photo.
// Needs Playwright with Chromium: node test/fixtures/build_pdf_sample.cjs
const {chromium}=require("playwright");
(async()=>{const b=await chromium.launch();const p=await b.newPage();
await p.setContent(`<html><body style="font-family:sans-serif;margin:0;padding:30px 40px">
<h1 style="color:#1e3a8a;margin:0 0 12px">Release Notes</h1>
<p style="font-size:14px;width:520px">This release makes sync <span style="color:#dc2626">twice as fast</span> and adds <span style="color:#16a34a;font-weight:bold">offline mode</span> for mobile users.</p>
<p style="font-size:14px;width:520px">See the <a href="https://example.com/docs">documentation site</a> for setup steps, or contact <a href="mailto:support@example.com">our support team</a>.</p>
<div style="display:flex;gap:20px;align-items:flex-start">
 <img id="photo" width="240" height="140">
 <p style="font-size:13px;width:260px">The chart on the left shows weekly active users growing by 40 percent after launch.</p>
</div>
<div style="position:relative;width:300px;height:120px;margin-top:16px">
 <img id="banner" width="300" height="120" style="position:absolute;left:0;top:0">
 <p style="position:absolute;left:16px;top:30px;margin:0;color:#fff;font-size:20px;font-weight:bold">Text over a photo</p>
</div>
</body></html>`);
await p.evaluate(()=>{
 const c=document.createElement("canvas");c.width=480;c.height=280;const x=c.getContext("2d");
 x.fillStyle="#f1f5f9";x.fillRect(0,0,480,280);x.fillStyle="#3b82f6";[60,90,120,170,230].forEach((h,i)=>x.fillRect(40+i*85,250-h,50,h));
 x.fillStyle="#111";x.font="24px sans-serif";x.fillText("Weekly users (image text)",40,40);
 document.getElementById("photo").src=c.toDataURL();
 const d=document.createElement("canvas");d.width=600;d.height=240;const y=d.getContext("2d");
 const g=y.createLinearGradient(0,0,600,240);g.addColorStop(0,"#7c3aed");g.addColorStop(1,"#f97316");y.fillStyle=g;y.fillRect(0,0,600,240);
 for(let i=0;i<40;i++){y.fillStyle=`rgba(255,255,255,${Math.random()*0.3})`;y.beginPath();y.arc(Math.random()*600,Math.random()*240,Math.random()*30,0,7);y.fill();}
 document.getElementById("banner").src=d.toDataURL();});
await p.waitForTimeout(300);
await p.pdf({path:require("path").join(__dirname, "pdf-styled-sample.pdf"),format:"A4",printBackground:true});await b.close();})();
