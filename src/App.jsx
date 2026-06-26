import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";
import {
  Plus, Pencil, Trash2, Upload, Download, Search, X, LayoutDashboard,
  Hammer, Ruler, Save, AlertTriangle,
} from "lucide-react";

/* ───────────────────────── Référentiel statuts ───────────────────────── */
const STATUTS = [
  { key: "En exécution",    color: "#2E7D32", soft: "#E6F4E8" },
  { key: "Attribué",        color: "#1565C0", soft: "#E4EEFB" },
  { key: "En adjudication", color: "#E65100", soft: "#FCEBDD" },
  { key: "Programmé",       color: "#00838F", soft: "#DDF1F3" },
  { key: "En publication",  color: "#EF9A00", soft: "#FDF1DA" },
  { key: "DCE en cours",    color: "#607D8B", soft: "#ECEFF1" },
  { key: "Infructueux",     color: "#C62828", soft: "#FBE5E5" },
  { key: "Tributaire",      color: "#90A4AE", soft: "#ECEFF1" },
];
const statutMeta = (k) => STATUTS.find((s) => s.key === k) || STATUTS[7];

const NAVY = "#0A2D5E";
const GOLD = "#F9A825";
const STORE_KEY = "aaslm_aoo_2026";
const HEADERS = ["N° AOO/CA","N° Marché/CA","Objet","Programme / Plan d'action","Date Ouverture des Plis","Estimation M.O (DH TTC)","Attributaire / Titulaire","Montant (DH TTC)","Date extrait de PV","Statut","Observations"];
const FIELDS  = ["noAOO","noMarche","objet","programme","dateOuverture","estimation","attributaire","montant","datePV","statut","observations"];

/* ───────────────────────── Helpers ───────────────────────── */
const uid  = () => Math.random().toString(36).slice(2, 9);
const mk   = (o) => ({ id: uid(), noMarche: "", dateOuverture: "", attributaire: "", montant: null, datePV: "", observations: "", ...o });
const fmtDH = (n) => (n === null || n === undefined || n === "" ? "—" : Number(n).toLocaleString("fr-FR") + " DH");
const fmtM  = (n) => (n / 1e6).toLocaleString("fr-FR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + " M";
const sum   = (rows, key) => rows.reduce((a, r) => a + (Number(r[key]) || 0), 0);
const countBy = (rows, s) => rows.filter((r) => r.statut === s).length;

function excelSerialToFr(v) {
  if (v instanceof Date) {
    return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;
  }
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;
  }
  return v ? String(v) : "";
}

function deriveStatut(obs, attributaire, montant) {
  const t = (obs || "").toLowerCase();
  if (t.includes("infructueux"))  return "Infructueux";
  if (t.includes("adjudication")) return "En adjudication";
  if (t.includes("tributaire"))   return "Tributaire";
  if (t.includes("dce"))          return "DCE en cours";
  if (t.includes("publication"))  return "En publication";
  if (t.includes("programm") || t.includes("ouverture")) return "Programmé";
  if (t.includes("exécution") || t.includes("execution")) return "En exécution";
  if (attributaire && montant)    return "Attribué";
  return "Programmé";
}

/* ───────────────────────── Données initiales ───────────────────────── */
const SEED = {
  etudes: [
    mk({ noAOO:"12/AOO/AASLM/2025", noMarche:"04/AASLM/2026", objet:"Mission d'assistance à maîtrise d'ouvrage – Suivi et contrôle des études et travaux d'aménagement du site de la lagune de Marchica", programme:"Projet transversal", dateOuverture:"29/01/2026", estimation:10144800, attributaire:"TME", montant:12468000, datePV:"12/02/2026", statut:"En exécution", observations:"Marché en cours d'exécution" }),
    mk({ noAOO:"01/AOO/AASLM/2026", noMarche:"03/AASLM/2026", objet:"Maîtrise d'œuvre – Requalification et réaménagement du Bd Mohammed V, placettes, espaces publics et réseaux viaires adjacents", programme:"Rénovation centre historique Damier : Bds Mohammed V et Hassan II", dateOuverture:"10/02/2026", estimation:1606320, attributaire:"SAFED", montant:1577880, datePV:"19/02/2026", statut:"En exécution", observations:"Marché en cours d'exécution" }),
    mk({ noAOO:"02/AOO/AASLM/2026", noMarche:"05/AASLM/2026", objet:"Étude et suivi des travaux de maintenance éclairage et mobilier – Zone Marchica", programme:"Mise à niveau des voiries (éclairage, signalisation, mobilier urbain…)", dateOuverture:"10/02/2026", estimation:840000, attributaire:"BTP CONSULTING", montant:798000, datePV:"18/02/2026", statut:"En exécution", observations:"Marché en cours d'exécution" }),
    mk({ noAOO:"04/AOO/AASLM/2026", noMarche:"08/AASLM/2026", objet:"Suivi des travaux de restructuration et requalification – Secteur Sidi Moussa, commune de Beni Ansar", programme:"Restructuration et mise à niveau des quartiers : Sidi Moussa – Beni Ansar", dateOuverture:"05/03/2026", estimation:396000, attributaire:"STE ESMOR", montant:271200, datePV:"13/03/2026", statut:"En exécution", observations:"Marché en cours d'exécution" }),
    mk({ noAOO:"05/AOO/AASLM/2026", noMarche:"11/AASLM/2026", objet:"Étude technique et suivi des travaux de protection de la ville de Nador contre les inondations de l'Oued Caballo", programme:"Ouvrages maritimes", dateOuverture:"07/04/2026", estimation:300000, attributaire:"HANDASSAT EL GHAD", montant:225000, datePV:"14/04/2026", statut:"Attribué" }),
    mk({ noAOO:"06/AOO/AASLM/2026", noMarche:"10/AASLM/2026", objet:"Suivi et contrôle extérieur topographique – Restructuration et mise à niveau des quartiers Sidi Moussa-Ghassi, Beni Ansar", programme:"Projet transversal", dateOuverture:"07/04/2026", estimation:58440, attributaire:"JTTOPO", montant:53160, datePV:"14/04/2026", statut:"Attribué" }),
    mk({ noAOO:"09/AOO/AASLM/2026", noMarche:"16/AASLM/2026", objet:"Suivi des travaux de finalisation du projet d'aménagement de Jnan El Matar – Secteur 4", programme:"Finalisation du Bd Jnan El Matar – Secteur 4", dateOuverture:"14/04/2026", estimation:228000, attributaire:"ESMOR", montant:148800, datePV:"27/04/2026", statut:"Attribué" }),
    mk({ noAOO:"10/AOO/AASLM/2026", noMarche:"14/AASLM/2026", objet:"Suivi et contrôle extérieur topographique des travaux de finalisation du projet d'aménagement de Jnan El Matar – Secteur 4", programme:"Finalisation du Bd Jnan El Matar – Secteur 4", dateOuverture:"14/04/2026", estimation:78000, attributaire:"JTTOPO", montant:59160, datePV:"21/04/2026", statut:"Attribué" }),
    mk({ noAOO:"15/AOO/AASLM/2026", noMarche:"17/AASLM/2026", objet:"Prestations de laboratoire – Contrôle extérieur des travaux et études géotechniques", programme:"Projet transversal", dateOuverture:"13/05/2026", estimation:1980540, attributaire:"LPEE", montant:2374920, datePV:"21/05/2026", statut:"Attribué" }),
    mk({ noAOO:"11/AOO/AASLM/2026", noMarche:"15/AASLM/2026", objet:"Étude de faisabilité – Mise en place d'infrastructures de débarquement des produits d'aquaculture au niveau de la lagune de Marchica", programme:"Études (Projet accompagnant les activités de pêche – aquaculture)", dateOuverture:"16/04/2026", estimation:780000, attributaire:"AL KHIBRA ÉTUDES ET CONSEILS", montant:585000, datePV:"22/04/2026", statut:"Attribué" }),
    mk({ noAOO:"24/AOO/AASLM/2026", objet:"Études techniques pour la maîtrise d'œuvre des projets d'aménagement de l'entrée Beni Ansar et création de la centralité jusqu'au point de passage", programme:"Aménagement entrée Beni Ansar et création de la centralité jusqu'au poste frontière", dateOuverture:"30/06/2026", estimation:3085200, statut:"Programmé", observations:"Ouverture des plis programmée le 30/06/2026" }),
    mk({ noAOO:"22/AOO/AASLM/2026", objet:"Étude et suivi des travaux d'équipement et de réalisation des ouvrages de protection et d'accostage du plan d'eau de la lagune de Marchica", programme:"Ouvrages maritimes", dateOuverture:"30/06/2026", estimation:1881000, statut:"Programmé", observations:"Ouverture des plis programmée le 30/06/2026" }),
    mk({ noAOO:"19/AOO/AASLM/2026", objet:"Études, suivi et contrôle extérieur topographique & topo-bathymétrique des travaux d'aménagement du site de la lagune de Marchica", programme:"Projet transversal", dateOuverture:"10/06/2026", estimation:1695000, statut:"Programmé", observations:"Ouverture programmée le 10/06/2026" }),
    mk({ noAOO:"—", objet:"Étude technique des travaux d'aménagement de la Zone Animation et loisirs Noor", programme:"Projet esplanade Noor et plages artificielles", estimation:null, statut:"Tributaire", observations:"Tributaire du lancement du marché global des zones paysagères" }),
  ],
  travaux: [
    mk({ noAOO:"03/AOO/AASLM/2025", noMarche:"01/AASLM/2026", objet:"Travaux de déviation d'un tronçon de ligne HTA souterrain traversant le projet Jnan El Matar Secteur 4", programme:"Finalisation du Bd Jnan El Matar – Secteur 4", dateOuverture:"28/07/2025", estimation:728167.2, attributaire:"SEDERAM", montant:750021.22, datePV:"01/08/2025", statut:"En exécution", observations:"Marché en cours d'exécution" }),
    mk({ noAOO:"09/AOO/AASLM/2025", noMarche:"02/AASLM/2026", objet:"Travaux d'électrification 2/1° catégorie du lotissement de recasement de Bouarourou", programme:"Finalisation du Bd Jnan El Matar – Secteur 4", dateOuverture:"13/11/2025", estimation:1730288.52, attributaire:"SEDERAM", montant:1903317.37, datePV:"13/11/2025", statut:"En exécution", observations:"Marché en cours d'exécution" }),
    mk({ noAOO:"03/AOO/AASLM/2026", noMarche:"07/AASLM/2026", objet:"Fourniture et installation du système VTS de la capitainerie de l'Agence Marchica", programme:"Gestion du plan d'eau", dateOuverture:"05/03/2026", estimation:923160, attributaire:"IT SOLUTIONS INOV", montant:837600, datePV:"13/03/2026", statut:"Attribué" }),
    mk({ noAOO:"07/AOO/AASLM/2026", noMarche:"09/AASLM/2026", objet:"Restructuration et mise à niveau des quartiers : Sidi Moussa-Ghassi – Tranche 1, commune de Beni Ansar", programme:"Restructuration et mise à niveau des quartiers : Sidi Moussa – Beni Ansar", dateOuverture:"07/04/2026", estimation:10216704, attributaire:"Groupement STE AFRICAINE DE TRAVAUX PUBLICS", montant:10100700, datePV:"14/04/2026", statut:"En exécution", observations:"Marché en cours d'exécution" }),
    mk({ noAOO:"13/AOO/AASLM/2026", noMarche:"12/AASLM/2026", objet:"Chargement, transport et étalage du sable au niveau des plages artificielles", programme:"Projet esplanade Noor et plages artificielles", dateOuverture:"28/04/2026", estimation:1056000, attributaire:"STE EXTRAT SERVICES D'AMÉNAGEMENT", montant:897600, datePV:"28/04/2026", statut:"Attribué" }),
    mk({ noAOO:"14/AOO/AASLM/2026", objet:"Travaux de finalisation d'aménagement du Bd Jnan El Matar Secteur 4 – Lot Revêtement", programme:"Finalisation du Bd Jnan El Matar – Secteur 4", dateOuverture:"28/04/2026", estimation:null, statut:"Infructueux", observations:"Infructueux – relancé sous N°16" }),
    mk({ noAOO:"16/AOO/AASLM/2026", objet:"Travaux de finalisation d'aménagement du Bd Jnan El Matar Secteur 4 – Lot Revêtement (relancé)", programme:"Finalisation du Bd Jnan El Matar – Secteur 4", dateOuverture:"02/06/2026", estimation:4262322, attributaire:"STE NEW WOY TRAV", montant:4968084, statut:"En adjudication", observations:"Marché en cours d'adjudication" }),
    mk({ noAOO:"Consultation BC", objet:"Travaux d'entretien du point de passage – commune de Beni Ansar", programme:"Mise à niveau des voiries (éclairage, signalisation, mobilier urbain…)", dateOuverture:"08/06/2026", estimation:282000, statut:"En exécution", observations:"En cours d'exécution" }),
    mk({ noAOO:"—", objet:"Réalisation des travaux préparatoires des projets d'aménagement de l'Agence", programme:"Mise à niveau des voiries (éclairage, signalisation, mobilier urbain…)", estimation:4119888, statut:"En publication", observations:"AO en cours de publication (envoyé au service achat le 26/05/2026)" }),
    mk({ noAOO:"23/AOO/AASLM/2026", objet:"Travaux de protection et d'équipement du cordon de Boukana", programme:"Protection et préservation des zones sensibles", estimation:2060400, statut:"En publication", observations:"AO en cours de publication (envoyé au service achat le 04/05/2026)" }),
    mk({ noAOO:"—", objet:"Mise en place d'éclairage et de système de surveillance de la capitainerie", programme:"Gestion du plan d'eau", estimation:2000000, statut:"DCE en cours", observations:"DCE en cours d'élaboration (AMO) – livrable prévu le 05/06/2026" }),
    mk({ noAOO:"—", objet:"Travaux de plantation de Jnan El Matar Secteur 4 et renforcement des espaces verts dans la zone d'intervention de l'Agence Marchica", programme:"Projet transversal", estimation:5332260, statut:"DCE en cours", observations:"DCE en cours de validation par la DG" }),
    mk({ noAOO:"23/AOO/AASLM/2026", objet:"Réalisation des travaux de signalisation verticale et horizontale des voies dans la zone d'intervention de l'Agence Marchica", programme:"Projet transversal", dateOuverture:"30/06/2026", estimation:1242420, statut:"Programmé", observations:"Ouverture des plis le 30/06/2026" }),
    mk({ noAOO:"—", objet:"Travaux de maintenance d'éclairage public de la zone Beni Ansar – Tranche 1, commune de Beni Ansar, province de Nador", programme:"Mise à niveau des voiries (éclairage, signalisation, mobilier urbain…)", estimation:1600000, statut:"DCE en cours", observations:"DCE en phase de finalisation" }),
  ],
};

/* ───────────────────────── UI atoms ───────────────────────── */
function StatutBadge({ statut }) {
  const m = statutMeta(statut);
  return (
    <span className="inline-flex items-center whitespace-nowrap rounded-full px-2.5 py-0.5 text-xs font-bold"
      style={{ background: m.soft, color: m.color }}>
      <span className="mr-1.5 h-1.5 w-1.5 rounded-full inline-block" style={{ background: m.color }} />
      {statut}
    </span>
  );
}

function Kpi({ value, label, color }) {
  return (
    <div className="flex flex-col justify-between rounded-xl p-4 text-white shadow-sm" style={{ background: color }}>
      <div className="text-3xl font-extrabold leading-none">{value}</div>
      <div className="mt-2 text-xs font-medium opacity-90">{label}</div>
    </div>
  );
}

/* ───────────────────────── Modal édition ───────────────────────── */
function EditModal({ row, onSave, onClose }) {
  const [f, setF] = useState(() => ({ ...row }));
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const num = (v) => (v === "" || v === null ? null : Number(String(v).replace(/\s/g, "").replace(",", ".")));

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-black/40 p-4" style={{paddingTop:"5vh"}} onClick={onClose}>
      <div className="w-full max-w-2xl rounded-2xl bg-white shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between rounded-t-2xl px-6 py-4" style={{ background: NAVY }}>
          <h3 className="text-base font-bold text-white">{row._isNew ? "Ajouter un appel d'offres" : "Modifier l'appel d'offres"}</h3>
          <button onClick={onClose} className="rounded-lg p-1 text-white/80 hover:text-white"><X size={20} /></button>
        </div>
        <div className="grid gap-4 overflow-y-auto p-6 grid-cols-2" style={{maxHeight:"65vh"}}>
          <Field label="N° AOO / CA"><input className="inp" value={f.noAOO||""} onChange={e=>set("noAOO",e.target.value)} placeholder="ex. 14/AOO/AASLM/2026" /></Field>
          <Field label="N° Marché / CA"><input className="inp" value={f.noMarche||""} onChange={e=>set("noMarche",e.target.value)} /></Field>
          <Field label="Objet" full><textarea className="inp h-20 resize-none" value={f.objet||""} onChange={e=>set("objet",e.target.value)} /></Field>
          <Field label="Programme / Plan d'action" full><input className="inp" value={f.programme||""} onChange={e=>set("programme",e.target.value)} /></Field>
          <Field label="Date d'ouverture des plis"><input className="inp" value={f.dateOuverture||""} onChange={e=>set("dateOuverture",e.target.value)} placeholder="JJ/MM/AAAA" /></Field>
          <Field label="Date extrait de PV"><input className="inp" value={f.datePV||""} onChange={e=>set("datePV",e.target.value)} placeholder="JJ/MM/AAAA" /></Field>
          <Field label="Estimation M.O (DH TTC)"><input className="inp text-right" value={f.estimation??""} onChange={e=>set("estimation",e.target.value)} /></Field>
          <Field label="Montant attribué (DH TTC)"><input className="inp text-right" value={f.montant??""} onChange={e=>set("montant",e.target.value)} /></Field>
          <Field label="Attributaire / Titulaire" full><input className="inp" value={f.attributaire||""} onChange={e=>set("attributaire",e.target.value)} /></Field>
          <Field label="Statut">
            <select className="inp" value={f.statut||"Programmé"} onChange={e=>set("statut",e.target.value)}>
              {STATUTS.map(s=><option key={s.key} value={s.key}>{s.key}</option>)}
            </select>
          </Field>
          <Field label="Observations" full><input className="inp" value={f.observations||""} onChange={e=>set("observations",e.target.value)} /></Field>
        </div>
        <div className="flex justify-end gap-2 rounded-b-2xl border-t bg-slate-50 px-6 py-4">
          <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-200">Annuler</button>
          <button onClick={()=>onSave({...f,estimation:num(f.estimation),montant:num(f.montant)})}
            className="flex items-center gap-2 rounded-lg px-4 py-2 text-sm font-bold text-white"
            style={{background:NAVY}}>
            <Save size={16}/> Enregistrer
          </button>
        </div>
      </div>
    </div>
  );
}
function Field({ label, children, full }) {
  return (
    <label className={`flex flex-col gap-1 ${full ? "col-span-2" : ""}`}>
      <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      {children}
    </label>
  );
}

/* ───────────────────────── Table ───────────────────────── */
function AoTable({ rows, onEdit, onDelete, accent }) {
  return (
    <div className="overflow-x-auto rounded-xl border border-slate-200 bg-white shadow-sm">
      <table className="w-full border-collapse text-sm">
        <thead>
          <tr style={{background:accent}} className="text-left text-white">
            {["N° AOO","Objet","Date ouv.","Estimation","Attributaire","Montant","Statut","Actions"].map(h=>(
              <th key={h} className={`px-3 py-2.5 text-xs font-semibold ${["Estimation","Montant"].includes(h)?"text-right":h==="Actions"?"text-center":""}`}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.length===0 && <tr><td colSpan={8} className="px-3 py-10 text-center text-slate-400">Aucun appel d'offres.</td></tr>}
          {rows.map((r,i)=>(
            <tr key={r.id} className={`border-t border-slate-100 align-top hover:bg-slate-50 ${i%2?"bg-slate-50/40":""}`}>
              <td className="px-3 py-2.5 font-bold" style={{color:NAVY}}>
                {r.noAOO}
                {r.noMarche&&<div className="mt-0.5 text-xs font-medium text-slate-400">Marché {r.noMarche}</div>}
              </td>
              <td className="max-w-sm px-3 py-2.5">
                <div className="text-slate-700" style={{display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical",overflow:"hidden"}}>{r.objet}</div>
                <div className="mt-0.5 text-xs text-slate-400">{r.programme}</div>
              </td>
              <td className="whitespace-nowrap px-3 py-2.5 text-slate-600">{r.dateOuverture||"—"}</td>
              <td className="whitespace-nowrap px-3 py-2.5 text-right text-slate-600">{fmtDH(r.estimation)}</td>
              <td className="px-3 py-2.5 text-slate-700">{r.attributaire||"—"}</td>
              <td className="whitespace-nowrap px-3 py-2.5 text-right font-semibold" style={{color:r.montant?statutMeta(r.statut).color:"#94A3B8"}}>{fmtDH(r.montant)}</td>
              <td className="px-3 py-2.5"><StatutBadge statut={r.statut}/></td>
              <td className="whitespace-nowrap px-3 py-2.5 text-center">
                <button onClick={()=>onEdit(r)} className="mr-1 rounded-lg p-1.5 text-slate-500 hover:text-blue-600" title="Modifier"><Pencil size={15}/></button>
                <button onClick={()=>onDelete(r)} className="rounded-lg p-1.5 text-slate-500 hover:text-red-600" title="Supprimer"><Trash2 size={15}/></button>
              </td>
            </tr>
          ))}
        </tbody>
        {rows.length>0&&(
          <tfoot>
            <tr style={{background:NAVY}} className="text-white">
              <td className="px-3 py-2.5 text-xs font-bold" colSpan={3}>TOTAL ({rows.length} AO)</td>
              <td className="whitespace-nowrap px-3 py-2.5 text-right text-sm font-bold">{fmtDH(sum(rows,"estimation"))}</td>
              <td/>
              <td className="whitespace-nowrap px-3 py-2.5 text-right text-sm font-bold">{fmtDH(sum(rows,"montant"))}</td>
              <td colSpan={2}/>
            </tr>
          </tfoot>
        )}
      </table>
    </div>
  );
}

/* ───────────────────────── Synthèse ───────────────────────── */
function Synthese({ data }) {
  const all=[...data.etudes,...data.travaux];
  const estT=sum(data.travaux,"estimation"),estE=sum(data.etudes,"estimation");
  const mtT=sum(data.travaux,"montant"),mtE=sum(data.etudes,"montant");
  const chartData=[
    {name:"Travaux",Estimation:estT,Attribué:mtT},
    {name:"Études & Services",Estimation:estE,Attribué:mtE},
  ];
  const breakdown=STATUTS.map(s=>({statut:s.key,color:s.color,n:countBy(all,s.key)})).filter(b=>b.n>0);
  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
        <Kpi value={data.travaux.length} label="AO Travaux (total)" color={NAVY}/>
        <Kpi value={data.etudes.length} label="AO Études & Services" color="#00838F"/>
        <Kpi value={countBy(all,"En exécution")} label="Marchés en exécution" color="#2E7D32"/>
        <Kpi value={countBy(all,"Attribué")} label="Marchés attribués" color="#1565C0"/>
      </div>
      <div className="grid grid-cols-1 gap-6 lg:grid-cols-5">
        <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm lg:col-span-3">
          <h3 className="mb-1 text-sm font-bold" style={{color:NAVY}}>Estimation vs Montant attribué</h3>
          <p className="mb-4 text-xs text-slate-400">Par catégorie (DH TTC)</p>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={chartData} margin={{top:8,right:8,bottom:0,left:8}}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF2F6" vertical={false}/>
              <XAxis dataKey="name" tick={{fontSize:12,fill:"#64748B"}}/>
              <YAxis tickFormatter={v=>(v/1e6).toFixed(0)+"M"} tick={{fontSize:11,fill:"#94A3B8"}}/>
              <Tooltip formatter={v=>Number(v).toLocaleString("fr-FR")+" DH"} contentStyle={{borderRadius:10,fontSize:12}}/>
              <Legend wrapperStyle={{fontSize:12}}/>
              <Bar dataKey="Estimation" fill="#1565C0" radius={[4,4,0,0]}/>
              <Bar dataKey="Attribué" fill="#00838F" radius={[4,4,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm lg:col-span-2">
          <h3 className="mb-4 text-sm font-bold" style={{color:NAVY}}>Récapitulatif financier</h3>
          {[{label:"Travaux",est:estT,mt:mtT},{label:"Études & Services",est:estE,mt:mtE}].map(r=>(
            <div key={r.label} className="mb-2.5 flex items-center justify-between">
              <span className="text-sm text-slate-600">{r.label}</span>
              <div className="text-right">
                <div className="text-sm font-semibold">{fmtM(r.est)} DH</div>
                <div className="text-xs text-slate-400">attr. {fmtM(r.mt)} DH</div>
              </div>
            </div>
          ))}
          <div className="mt-3 border-t pt-3">
            <div className="flex items-center justify-between">
              <span className="text-sm font-bold" style={{color:NAVY}}>Total général estimé</span>
              <span className="text-lg font-extrabold" style={{color:GOLD}}>{fmtM(estT+estE)} DH</span>
            </div>
            <div className="mt-1 flex items-center justify-between text-xs text-slate-400">
              <span>Total attribué</span><span className="font-semibold">{fmtM(mtT+mtE)} DH</span>
            </div>
          </div>
          <div className="mt-5">
            <h4 className="mb-3 text-xs font-bold uppercase tracking-wide text-slate-500">Par statut</h4>
            <div className="space-y-1.5">
              {breakdown.map(b=>(
                <div key={b.statut} className="flex items-center justify-between rounded-lg border border-slate-100 px-3 py-1.5">
                  <div className="flex items-center gap-2">
                    <span className="h-2.5 w-2.5 rounded-full inline-block" style={{background:b.color}}/>
                    <span className="text-xs text-slate-600">{b.statut}</span>
                  </div>
                  <span className="text-sm font-bold" style={{color:b.color}}>{b.n}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ───────────────────────── App ───────────────────────── */
export default function App() {
  const [data, setData] = useState(SEED);
  const [tab, setTab] = useState("synthese");
  const [editing, setEditing] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [q, setQ] = useState("");
  const [fStatut, setFStatut] = useState("");
  const [toast, setToast] = useState("");
  const fileRef = useRef(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORE_KEY);
      if (saved) setData(JSON.parse(saved));
    } catch(e) {}
  }, []);
  useEffect(() => {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(data)); } catch(e) {}
  }, [data]);

  const flash = (m) => { setToast(m); setTimeout(()=>setToast(""),2600); };
  const cat = tab==="etudes"?"etudes":tab==="travaux"?"travaux":null;

  const filtered = useMemo(()=>{
    if(!cat) return [];
    return data[cat].filter(r=>{
      const okQ=!q||[r.noAOO,r.noMarche,r.objet,r.programme,r.attributaire,r.observations].join(" ").toLowerCase().includes(q.toLowerCase());
      const okS=!fStatut||r.statut===fStatut;
      return okQ&&okS;
    });
  },[data,cat,q,fStatut]);

  const save = (row)=>{
    const {_isNew,_cat,...clean}=row;
    setData(d=>{
      const list=d[_cat];
      const exists=list.some(x=>x.id===clean.id);
      return {...d,[_cat]:exists?list.map(x=>x.id===clean.id?clean:x):[...list,clean]};
    });
    setEditing(null);
    flash(_isNew?"Appel d'offres ajouté":"Modifications enregistrées");
  };
  const del = (row)=>{
    setData(d=>({...d,[cat]:d[cat].filter(x=>x.id!==row.id)}));
    setConfirmDel(null);
    flash("Appel d'offres supprimé");
  };
  const addNew = ()=>setEditing({...mk({statut:"Programmé"}),"_isNew":true,"_cat":cat});

  const onImport = async(e)=>{
    const file=e.target.files?.[0]; if(!file) return;
    try {
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array",cellDates:true});
      const parsed={etudes:[],travaux:[]};
      wb.SheetNames.forEach(sn=>{
        const isTrav=/trav/i.test(sn),isEtud=/etud|étud|service/i.test(sn);
        if(!isTrav&&!isEtud) return;
        const aoa=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:false,defval:""});
        const hi=aoa.findIndex(row=>row.some(c=>String(c).replace(/\s/g,"").toUpperCase().includes("N°AOO")));
        if(hi<0) return;
        for(let i=hi+1;i<aoa.length;i++){
          const row=aoa[i];
          const objet=String(row[2]||"").trim();
          if(!objet||/^total/i.test(String(row[0]||""))) continue;
          const est=parseNum(row[5]),mt=parseNum(row[7]),obs=String(row[11]||"").trim(),attr=String(row[6]||"").trim();
          (isTrav?parsed.travaux:parsed.etudes).push(mk({
            noAOO:String(row[0]||"").trim(),noMarche:String(row[1]||"").trim(),objet,
            programme:String(row[3]||"").trim(),dateOuverture:excelSerialToFr(row[4]),
            estimation:est,attributaire:attr,montant:mt,datePV:excelSerialToFr(row[8]),
            statut:deriveStatut(obs,attr,mt),observations:obs,
          }));
        }
      });
      if(parsed.etudes.length+parsed.travaux.length===0){flash("Aucune donnée reconnue");return;}
      setData(parsed);
      flash(`Importé : ${parsed.travaux.length} travaux, ${parsed.etudes.length} études`);
    } catch(err){ flash("Erreur de lecture du fichier"); }
    finally { if(fileRef.current) fileRef.current.value=""; }
  };
  const parseNum=(v)=>{if(v===""||v===null||v===undefined) return null; const n=Number(String(v).replace(/[\s ]/g,"").replace(/,/g,".")); return isNaN(n)?null:n;};

  const onExport = ()=>{
    const wb=XLSX.utils.book_new();
    const buildSheet=(rows,title)=>{
      const body=rows.map(r=>FIELDS.map(f=>r[f]??"")),aoa=[[title],HEADERS,...body,["TOTAL","","","","",sum(rows,"estimation"),"",sum(rows,"montant"),"","",""]];
      const ws=XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"]=[{wch:16},{wch:13},{wch:46},{wch:30},{wch:13},{wch:16},{wch:26},{wch:16},{wch:13},{wch:15},{wch:36}];
      ws["!merges"]=[{s:{r:0,c:0},e:{r:0,c:10}}];
      return ws;
    };
    const estT=sum(data.travaux,"estimation"),estE=sum(data.etudes,"estimation"),mtT=sum(data.travaux,"montant"),mtE=sum(data.etudes,"montant");
    const synth=[["SITUATION DES APPELS D'OFFRES — AASLM 2026"],["Mise à jour : "+new Date().toLocaleDateString("fr-FR")],[],["Indicateur","Valeur"],["AO Travaux",data.travaux.length],["AO Études & Services",data.etudes.length],["En exécution",countBy([...data.travaux,...data.etudes],"En exécution")],[],["Catégorie","Estimation","Attribué"],["Travaux",estT,mtT],["Études & Services",estE,mtE],["TOTAL GÉNÉRAL",estT+estE,mtT+mtE]];
    const wsS=XLSX.utils.aoa_to_sheet(synth);
    wsS["!cols"]=[{wch:32},{wch:18},{wch:18}];
    wsS["!merges"]=[{s:{r:0,c:0},e:{r:0,c:2}}];
    XLSX.utils.book_append_sheet(wb,wsS,"Synthèse");
    XLSX.utils.book_append_sheet(wb,buildSheet(data.etudes,"Études & Services 2026"),"Études & Services 2026");
    XLSX.utils.book_append_sheet(wb,buildSheet(data.travaux,"Travaux 2026"),"Travaux 2026");
    XLSX.writeFile(wb,"SITUATION_AOO_AASLM_2026.xlsx");
    flash("Fichier Excel exporté");
  };

  const tabs=[
    {key:"synthese",label:"Synthèse",icon:LayoutDashboard},
    {key:"travaux",label:"Travaux",icon:Hammer,n:data.travaux.length},
    {key:"etudes",label:"Études & Services",icon:Ruler,n:data.etudes.length},
  ];
  const accent=tab==="travaux"?"#123E73":"#00838F";

  return (
    <div className="min-h-screen bg-slate-100 font-sans text-slate-800">
      <style>{`.inp{border:1px solid #CBD5E1;border-radius:8px;padding:8px 10px;font-size:13px;outline:none;width:100%;background:#fff}.inp:focus{border-color:${NAVY};box-shadow:0 0 0 3px rgba(10,45,94,.12)}`}</style>
      <header className="text-white shadow-lg" style={{background:NAVY}}>
        <div className="h-1" style={{background:GOLD}}/>
        <div className="mx-auto flex max-w-7xl flex-col gap-3 px-4 py-4 sm:flex-row sm:items-center sm:justify-between sm:px-6">
          <div>
            <div className="text-xs font-bold uppercase tracking-widest" style={{color:GOLD}}>AASLM · Programme d'Action 2026</div>
            <h1 className="text-lg font-extrabold sm:text-xl">Situation des Appels d'Offres</h1>
          </div>
          <div className="flex gap-2">
            <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={onImport}/>
            <button onClick={()=>fileRef.current?.click()} className="flex items-center gap-2 rounded-lg bg-white/10 px-3 py-2 text-sm font-semibold text-white hover:bg-white/20">
              <Upload size={16}/> Importer Excel
            </button>
            <button onClick={onExport} className="flex items-center gap-2 rounded-lg px-3 py-2 text-sm font-bold text-slate-900" style={{background:GOLD}}>
              <Download size={16}/> Exporter Excel
            </button>
          </div>
        </div>
        <div className="mx-auto flex max-w-7xl gap-1 px-4 sm:px-6">
          {tabs.map(t=>{
            const Icon=t.icon,active=tab===t.key;
            return (
              <button key={t.key} onClick={()=>{setTab(t.key);setQ("");setFStatut("");}}
                className={`flex items-center gap-2 rounded-t-lg px-4 py-2.5 text-sm font-semibold ${active?"bg-slate-100 text-slate-900":"text-white/70 hover:bg-white/10 hover:text-white"}`}>
                <Icon size={16}/> {t.label}
                {t.n!==undefined&&<span className={`rounded-full px-1.5 py-0.5 text-xs font-bold ${active?"bg-slate-200 text-slate-700":"bg-white/20"}`}>{t.n}</span>}
              </button>
            );
          })}
        </div>
      </header>
      <main className="mx-auto max-w-7xl px-4 py-6 sm:px-6">
        {tab==="synthese" ? <Synthese data={data}/> : (
          <div className="space-y-4">
            <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex flex-1 flex-col gap-2 sm:flex-row sm:items-center">
                <div className="relative flex-1 sm:max-w-xs">
                  <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"/>
                  <input className="inp pl-9" placeholder="Rechercher…" value={q} onChange={e=>setQ(e.target.value)}/>
                </div>
                <select className="inp sm:w-52" value={fStatut} onChange={e=>setFStatut(e.target.value)}>
                  <option value="">Tous les statuts</option>
                  {STATUTS.map(s=><option key={s.key} value={s.key}>{s.key}</option>)}
                </select>
              </div>
              <button onClick={addNew} className="flex items-center justify-center gap-2 rounded-lg px-4 py-2 text-sm font-bold text-white" style={{background:NAVY}}>
                <Plus size={16}/> Ajouter un AO
              </button>
            </div>
            <AoTable rows={filtered} accent={accent} onEdit={r=>setEditing({...r,_cat:cat})} onDelete={setConfirmDel}/>
            <p className="text-xs text-slate-400">{filtered.length} appel(s) d'offres{q||fStatut?` (sur ${data[cat].length})`:""}</p>
          </div>
        )}
      </main>
      {editing && <EditModal row={editing} onClose={()=>setEditing(null)} onSave={save}/>}
      {confirmDel && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={()=>setConfirmDel(null)}>
          <div className="w-full max-w-sm rounded-2xl bg-white p-6 shadow-2xl" onClick={e=>e.stopPropagation()}>
            <div className="mb-3 flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-full bg-red-50 text-red-600"><AlertTriangle size={20}/></div>
              <h3 className="text-base font-bold">Supprimer cet AO ?</h3>
            </div>
            <p className="mb-5 text-sm text-slate-500">« {confirmDel.objet?.slice(0,80)}… » sera définitivement retiré.</p>
            <div className="flex justify-end gap-2">
              <button onClick={()=>setConfirmDel(null)} className="rounded-lg px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-100">Annuler</button>
              <button onClick={()=>del(confirmDel)} className="rounded-lg bg-red-600 px-4 py-2 text-sm font-bold text-white hover:bg-red-700">Supprimer</button>
            </div>
          </div>
        </div>
      )}
      {toast && (
        <div className="fixed bottom-5 left-1/2 z-50 -translate-x-1/2 rounded-lg px-4 py-2.5 text-sm font-bold text-white shadow-xl" style={{background:NAVY}}>
          {toast}
        </div>
      )}
    </div>
  );
}
