
import {validateLicense,readExcel,getSheet,addRecord,updateRecord,saveExcel} from "../excel.js";
const cfg=window.MODULE_CONFIG, $=id=>document.getElementById(id);
const session=JSON.parse(sessionStorage.getItem("accsystem_session")||"null");
if(!session){window.top.location.href="../login.html";throw new Error("No active session");}
const access=session.Access||[],admin=session.Role==="Administrator";
const allowed=perm=>admin||access.some(x=>String(x.Module).trim()===cfg.permission && (x[perm]===true||String(x[perm]).toLowerCase()==="true"||String(x[perm])==="1"));
if(!allowed("CanView")){document.body.innerHTML="<h2>Access denied</h2>";throw new Error("No view permission");}
let data=[];
const esc=v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const n=v=>Number(v)||0, now=()=>new Date().toISOString(), id=(prefix,arr,key)=>prefix+String(Math.max(0,...arr.map(x=>parseInt(String(x[key]||"").replace(/\D/g,""))||0))+1).padStart(5,"0");
const sheet=s=>{try{return getSheet(s)||[]}catch{return []}};
function message(t,bad=false){$("msg").textContent=t;$("msg").className="msg "+(bad?"bad":"good");}
function options(field){if(!field.options)return "";let vals=field.options==="customers"?sheet("Customers"):field.options==="vendors"?sheet("Vendors"):field.options==="items"?sheet("Items"):field.options==="banks"?sheet("Banks"):field.options==="cash"?sheet("CashAccounts"):field.options==="coa"?sheet("ChartOfAccounts"):[];return vals.map(x=>{let value=x[field.valueKey||"CustomerID"];let label=x[field.labelKey||"CustomerName"]||x.AccountName||x.ItemName||x.BankName||x.CashAccountName||value;return `<option value="${esc(value)}">${esc(label)}</option>`}).join("")}
function input(f){let v=f.value||"";if(f.type==="select")return `<select name="${f.name}" ${f.required?"required":""}>${f.options?options(f):f.choices.map(x=>`<option>${esc(x)}</option>`).join("")}</select>`;return `<input name="${f.name}" type="${f.type||"text"}" value="${esc(v)}" ${f.required?"required":""} ${f.step?`step="${f.step}"`:""} ${f.min!==undefined?`min="${f.min}"`:""}>`}
function render(){
 $("title").textContent=cfg.title;$("desc").textContent=cfg.desc;
 $("formFields").innerHTML=cfg.fields.map(f=>`<label>${esc(f.label)}${input(f)}</label>`).join("");
 $("save").disabled=!allowed("CanCreate");
 $("tableHead").innerHTML=cfg.columns.map(c=>`<th>${esc(c.label)}</th>`).join("")+"<th>Action</th>";
 $("tableBody").innerHTML=data.map((r,i)=>`<tr>${cfg.columns.map(c=>`<td>${esc(r[c.key])}</td>`).join("")}<td>${allowed("CanEdit")?`<button data-edit="${i}">Edit</button>`:""} ${allowed("CanDelete")?`<button class="danger" data-del="${i}">Deactivate</button>`:""}</td></tr>`).join("")||`<tr><td colspan="${cfg.columns.length+1}" class="empty">No records found.</td></tr>`;
}
async function load(){try{const lic=await validateLicense();if(lic===false||lic?.valid===false)throw new Error(lic.message||"License invalid");await readExcel();data=sheet(cfg.sheet);render();}catch(e){message("Database error: "+e.message,true)}}
$("form").addEventListener("submit",async e=>{e.preventDefault();try{if(!allowed("CanCreate"))throw Error("No create permission");const fd=new FormData(e.target), rec={};cfg.fields.forEach(f=>rec[f.name]=fd.get(f.name));cfg.fields.forEach(f=>{if(f.type==="number")rec[f.name]=n(rec[f.name])});const t=now();if(cfg.beforeSave)await cfg.beforeSave(rec);rec[cfg.key]=id(cfg.prefix,data,cfg.key);const docNo={Invoices:"InvoiceNo",Bills:"BillNo",CustomerPayments:"PaymentNo",VendorPayments:"PaymentNo"}[cfg.sheet];if(docNo&&!rec[docNo])rec[docNo]=id(cfg.prefix,data,docNo);rec.CreatedAt=t;rec.UpdatedAt=t;if(cfg.defaults)Object.assign(rec,typeof cfg.defaults==="function"?cfg.defaults(rec):cfg.defaults);addRecord(cfg.sheet,rec);if(cfg.afterSave)await cfg.afterSave(rec);await saveExcel();data=sheet(cfg.sheet);e.target.reset();render();message("Record saved successfully.");}catch(err){message(err.message,true)}});
$("tableBody").addEventListener("click",async e=>{const del=e.target.dataset.del,edit=e.target.dataset.edit;if(del!==undefined){const r=data[Number(del)];if(!confirm("Deactivate this record?"))return;try{if(!allowed("CanDelete"))throw Error("No delete permission");updateRecord(cfg.sheet,cfg.key,r[cfg.key],{IsActive:false,UpdatedAt:now()});await saveExcel();data=sheet(cfg.sheet);render();message("Record deactivated.");}catch(err){message(err.message,true)}}else if(edit!==undefined){const r=data[Number(edit)];if(!allowed("CanEdit"))return;const changes={};for(const f of cfg.fields){const v=prompt("Edit "+f.label,r[f.name]??"");if(v===null)return;changes[f.name]=f.type==="number"?n(v):v;}try{updateRecord(cfg.sheet,cfg.key,r[cfg.key],{...changes,UpdatedAt:now()});await saveExcel();data=sheet(cfg.sheet);render();message("Record updated.");}catch(err){message(err.message,true)}}});
load();
